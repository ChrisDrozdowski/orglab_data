#include <iostream>
#include <conio.h>
#include <exception>
#include <vector>
#include <random>
#include <chrono>
#include <mutex>
#include <filesystem>
#include <Shlobj.h>
#include <atlsafe.h>

// #define this to use orglab rather than Origin automation Server.
//#define ORGLAB_DATA_USE_ORGLAB


// ORGLAB_DATA_USE_ORGLAB may or may not be #define'd at top of file.
#ifdef ORGLAB_DATA_USE_ORGLAB
// Import the Orglab type library and assign it to the origin namespace.
#import "orglab.tlb" no_dual_interfaces rename_namespace("origin")
#else
// Import the Origin Automation Server type library and assign it to the origin namespace.
#import "Origin8.tlb" no_dual_interfaces rename_namespace("origin")
#endif

// You MUST define this prior to including orglab_data.hpp.
// It should be set to the namespace you use for the type library.
// For this example, the namespace is origin (see above).
#define ORGLAB_DATA_ORGLAB_NS origin

// Include orglab_data.hpp after importing type library and other includes.
#include "../orglab_data.hpp"


namespace my_utils {
	template <typename T>
	std::vector<T> get_test_data(long n, long initial = 0, bool sort = false) {
		std::random_device rd;
		std::default_random_engine rng(rd());
		std::uniform_real_distribution<> dist(0, n);
		std::vector<T> data;
		data.reserve(n);
		std::generate_n(std::back_inserter(data), n, [&]() {
			return initial + dist(rng);
			});
		if (sort)
			std::sort(data.begin(), data.end());
		return data;
	}

	std::chrono::time_point<std::chrono::steady_clock> start;
	std::chrono::milliseconds elapsed_ms(bool init = false) {
		if (init) start = std::chrono::steady_clock::now();
		return std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start);
	}

	std::wstring get_user_documents_folder() {
		std::wstring wstr;
		PWSTR buf1 = NULL;
		if (S_OK == ::SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &buf1))// Return long path.
			wstr = buf1;
		::CoTaskMemFree(buf1);
		if (0 == wstr.length())
			return std::wstring{};
		std::filesystem::path path(wstr);
		std::error_code ec;
		if (!std::filesystem::is_directory(path, ec))
			return std::wstring{};
		return wstr + std::filesystem::path::preferred_separator;
	}

}

// Custom function replaces default function used for COM error handling.
// Converts COM error to std::exception. Very useful!
void __stdcall my_com_raise_error(HRESULT hr, IErrorInfo* perrorinfo) {
	if (perrorinfo) {
		BSTR bstr;
		perrorinfo->GetDescription(&bstr);
		char* ch = ::_com_util::ConvertBSTRToString(bstr);
		::SysFreeString(bstr);
		std::string str(ch);
		delete[] ch;
		throw std::exception(str.c_str());
	}
	else
		throw std::exception("Unknown COM Error");
}

// Ensures that above function is called only one time.
void my_set_com_error_handler() {
	static std::once_flag flag;
	std::call_once(flag, []() {
		::_set_com_error_handler(my_com_raise_error);
		});
}


int main()
{
	// Install custom function defined above used for COM error handling.
	// This is not documented in OriginLab data but I discovered it.
	my_set_com_error_handler();

	// Start initialization of COM.
	::CoInitializeEx(nullptr, COINIT_DISABLE_OLE1DDE | COINIT_MULTITHREADED);
	CLSID clsid;
	CComPtr<origin::IOApplication> app;

// ORGLAB_DATA_USE_ORGLAB may or may not be #define'd at top of file.
#ifdef ORGLAB_DATA_USE_ORGLAB
	// This is the launch sequence for Orglab.
	const std::vector<std::wstring> prog_ids = { L"OrgLab995.Application", L"OrgLab990.Application", L"OrgLab99.Application",
	L"OrgLab985.Application", L"OrgLab980.Application", L"OrgLab98.Application",
	L"OrgLab975.Application", L"OrgLab970.Application", L"OrgLab97.Application", L"OrgLab9.Application"
	};
	for (const std::wstring str : prog_ids) {
		if (SUCCEEDED(::CLSIDFromProgID(str.c_str(), &clsid)))
			break;
	}
	app.CoCreateInstance(clsid);
#else
	// This is the launch sequence for Origin Automation Server.
	::CLSIDFromProgID(L"Origin.Application", &clsid);
	::CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&app);
	app->Visible = origin::MAINWND_SHOW;
#endif

	{ // Scope ensures COM pointers released properly. Important!

		// Begin examples in this scope.

		/*
		Mapping of supported C++ data type to supported Origin data types and vice versa.
		Only these C++ data types are supported.
		double					<==>	COLDATAFORMAT::DF_TEXT_NUMERIC
		double					<==>	COLDATAFORMAT::DF_DOUBLE
		double					<==>	COLDATAFORMAT::DF_DATE
		double					<==>	COLDATAFORMAT::DF_TIME
		float					<==>	COLDATAFORMAT::DF_FLOAT
		int						<==>	COLDATAFORMAT::DF_LONG
		long					<==>	COLDATAFORMAT::DF_LONG
		unsigned long			<==>	COLDATAFORMAT::DF_ULONG
		short					<==>	COLDATAFORMAT::DF_SHORT
		unsigned short			<==>	COLDATAFORMAT::DF_USHORT
		std::wstring			<==>	COLDATAFORMAT::DF_TEXT_NUMERIC
		std::string				<==>	COLDATAFORMAT::DF_TEXT_NUMERIC
		byte					<==>	COLDATAFORMAT::DF_BYTE
		char					<==>	COLDATAFORMAT::DF_CHAR
		std::complex<double>	<==>	COLDATAFORMAT::DF_COMPLEX
		*/

		// Setting column data.

		// Create a new workbook with 5 columns and assign column pointers.
		origin::WorksheetPagePtr wksp_ptr = app->WorksheetPages->Add();
		origin::WorksheetPtr wks_ptr = wksp_ptr->Layers->Item[0];
		wks_ptr->Cols = 5;
		wks_ptr->Labels(L"LUC");// Make sure labels rows get displayed. May be flakey.
		origin::ColumnPtr col_ptr_1 = wks_ptr->Columns->Item[0];
		origin::ColumnPtr col_ptr_2 = wks_ptr->Columns->Item[1];
		origin::ColumnPtr col_ptr_3 = wks_ptr->Columns->Item[2];
		origin::ColumnPtr col_ptr_4 = wks_ptr->Columns->Item[3];
		origin::ColumnPtr col_ptr_5 = wks_ptr->Columns->Item[4];


		// Vector and array functions:
		// void set_column_data(const ColumnPtr& ptr, const std::vector<T>& data, const std::size_t& offset = 0)
		// void set_column_data(const ColumnPtr& ptr, const T* data, const std::size_t& rows, const std::size_t& offset = 0)
		// Array version only for numeric and complex- not string data.

		// When setting column data, the column's data type will be changed to match
		// the first one that maps to the C++ data type. This can be turned off by:
		// #define ORGLAB_DATA_NO_CHANGE_DATA_TYPE

		// Exception thrown if ColumnPtr is invalid or can't set data for some reason.

		std::vector<double> vec_1 = { 1,2,2.3,3.4,4.5,5.6 };
		orglab_data::set_column_data(col_ptr_1, vec_1);

		std::vector<unsigned long> vec_2 = { 1, 2, 3, 4, 5 };
		orglab_data::set_column_data(col_ptr_2, vec_2);

		// std::wstring.
		std::vector<std::wstring> vec_3 = { L"hello world", L"مرحبا بالعالم", L"Բարեւ աշխարհ",
			L"Здравей свят", L"Прывітанне Сусвет", L"မင်္ဂလာပါကမ္ဘာလောက", L"你好，世界",
			L"Γειά σου Κόσμε", L"હેલ્લો વિશ્વ", L"Helló Világ", L"こんにちは世界", L"안녕 세상",
			L"سلام دنیا", L"העלא וועלט" };
		orglab_data::set_column_data(col_ptr_3, vec_3);

		// std::string.
		// Below observe data starts at the 14th (zero-based) row in column (offset param = 14).
		// Much slower than wide strings because it must convert from string to wide string.
		std::vector<std::string> vec_4 = { "Simple string", "Another simple string" };
		orglab_data::set_column_data(col_ptr_3, vec_4, 14);

		// Array version.
		unsigned short arr[5] = { 123, 234, 345, 456, 567 };
		orglab_data::set_column_data(col_ptr_4, arr, 5);

		// Complex data.
		std::vector<std::complex<double>> vec_5 = { std::complex<double>(1, 2),
			std::complex<double>(3, 4), std::complex<double>(5, 6) };
		orglab_data::set_column_data(col_ptr_5, vec_5);


		// Getting column data.

		// Functions:
		// std::vector<T> get_column_data<T>(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1)

		// When getting data from a column, you must assign it to a vector whose
		// C++ data type is mapped to the relevant column data type mapped above.
		// Otherwise, an exception is thrown.
		// Observe how type hints are used.

		// Exception thrown if ColumnPtr is invalid or an incompatible data type is specified.

		std::vector<double> vec_6 = orglab_data::get_column_data<double>(col_ptr_1);
		std::vector<unsigned long> vec_7 = orglab_data::get_column_data<unsigned long>(col_ptr_2);

		// Below observe offset & number of rows to return.
		std::vector<std::wstring> vec_8 = orglab_data::get_column_data<std::wstring>(col_ptr_3, 2, 3);
		std::vector<std::string> vec_9 = orglab_data::get_column_data<std::string>(col_ptr_3, 4, 5);

		// In this case, an exception will be throw because the specified data type does not match the column data type.
		try {
			std::vector<int> vec_10 = orglab_data::get_column_data<int>(col_ptr_1);
		}
		catch (const std::exception & e) {
			std::string s = e.what();
		}


		// How about a performance test?!?
		wks_ptr->Cols = wks_ptr->Cols++;
		origin::ColumnPtr col_ptr_1e6 = wks_ptr->Columns->Item[wks_ptr->Cols - 1];
		std::vector<double> vec_in_1e6 = my_utils::get_test_data<double>(1e6);

		my_utils::elapsed_ms(true);
		orglab_data::set_column_data<double>(col_ptr_1e6, vec_in_1e6);
		std::cout << "Write 1E6 rows: " << my_utils::elapsed_ms().count() << " ms" << std::endl;
		my_utils::elapsed_ms(true);
		std::vector<double> vec_out_1e6 = orglab_data::get_column_data<double>(col_ptr_1e6);
		std::cout << "Read 1E6 rows: " << my_utils::elapsed_ms().count() << " ms" << std::endl;


		// Setting and getting matrix data.

		// Functions:
		// void set_matrix_data(const MatrixObjectPtr& ptr, const orglab_data::matrix_adapter<T>& ma)
		// orglab_data::matrix_adapter<T> = get_matrix_data(const MatrixObjectPtr& ptr)
		// Same rules apply to data types and exceptions as those for columns.

		// About orglab_data::matrix_adapter class.
		// Simple 2D C++ "matrix" class meant to make it
		// easier to interact with MatrixObjectPtr.
		// See beginning of orglab_data::impl namespace for the class definition.

		// Create a new matrixbook with 3 matrices and assign matrix pointers.
		origin::MatrixPagePtr mksp_ptr = app->MatrixPages->Add();
		origin::MatrixSheetPtr mks_ptr = mksp_ptr->Layers->Item[0];
		mks_ptr->Mats = 3;
		mks_ptr->Rows = 5; // VERY important. matrix dimensions must match this or exception.
		mks_ptr->Cols = 7; // VERY important. matrix dimensions must match this or exception.
		origin::MatrixObjectPtr mat_ptr_1 = mks_ptr->MatrixObjects->Item[0];
		origin::MatrixObjectPtr mat_ptr_2 = mks_ptr->MatrixObjects->Item[1];
		origin::MatrixObjectPtr mat_ptr_3 = mks_ptr->MatrixObjects->Item[2];

		unsigned short l;

		// Set double matrix.
		orglab_data::matrix_adapter<double> ma_1(5, 7); // 5 rows, 7 cols
		l = 1;
		for (unsigned short i = 0; i < 5; ++i) {
			for (unsigned short j = 0; j < 7; ++j, ++l) {
				ma_1(i, j) = 1.0 * l;// rows, cols.
			}
		}
		orglab_data::set_matrix_data(mat_ptr_1, ma_1);

		// Set unsigned short matrix.
		orglab_data::matrix_adapter<unsigned short> ma_2(5, 7); // 5 rows, 7 cols
		l = 1;
		for (unsigned short i = 0; i < 5; ++i) {
			for (unsigned short j = 0; j < 7; ++j, ++l) {
				ma_2(i, j) = 1 * l;// rows, cols.
			}
		}
		orglab_data::set_matrix_data(mat_ptr_2, ma_2);

		// Set complex<double> matrix.
		orglab_data::matrix_adapter<std::complex<double>> ma_3(5, 7); // 5 rows, 7 cols
		l = 1;
		for (unsigned short i = 0; i < 5; ++i) {
			for (unsigned short j = 0; j < 7; ++j, ++l) {
				std::complex<double> cpx{ 1.0 * l, 0.1 * l };
				ma_3(i, j) = cpx; // rows, cols.
			}
		}
		orglab_data::set_matrix_data(mat_ptr_3, ma_3);

		// Get double matrix.
		orglab_data::matrix_adapter<double> ma_4 = orglab_data::get_matrix_data<double>(mat_ptr_1);

		// Get unsigned short matrix.
		orglab_data::matrix_adapter<unsigned short> ma_5 = orglab_data::get_matrix_data<unsigned short>(mat_ptr_2);

		// Get complex<double> matrix.
		orglab_data::matrix_adapter<std::complex<double>> ma_6 = orglab_data::get_matrix_data<std::complex<double>>(mat_ptr_3);

		// You can dump matrices. orglab_data::matrix_adapter supports the << operator.
		std::cout << ma_6 << std::endl;


		// Now for string property handling.
		// Anywhere a string-based property (e.g. long name) needs to be written or read, these functions
		// make it easier.

		// Functions:
		// _bstr_t to_str_prop(const std::wstring& str)
		// _bstr_t to_str_prop(const std::string& str)
		// std::wstring from_str_prop<std::wstring>(const _bstr_t& prop)
		// std::string from_str_prop<std::string>(const _bstr_t& prop)

		// Set column long name property based on std::string and unit property based on std::wstring.
		col_ptr_1->LongName = orglab_data::to_str_prop("Time"); // std::string.
		col_ptr_1->Units = orglab_data::to_str_prop(L"sec"); // std::wstring.

		// Get the properties back. Observe we can get them back as std::string or std::wstring.
		std::wstring ln = orglab_data::from_str_prop<std::wstring>(col_ptr_1->LongName);
		std::string u = orglab_data::from_str_prop<std::string>(col_ptr_1->Units);


		// Finally, save the project to users Documents folder.
		std::wstring w_file = my_utils::get_user_documents_folder() + L"example.opju";
		_variant_t v_file(w_file.c_str());
		app->Save(v_file);
	}


	// This is same shut down sequence for Origin Automation Server and Orglab.
	app->Exit();
	app.Release();
	app = nullptr;
	::Sleep(500); // Apparently need a small delay to avoid an exception.
	::CoUninitialize(); // Must be called from same thread as ::CoInitialize.

}
