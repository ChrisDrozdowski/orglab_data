#include <iostream>
#include <conio.h>
#include <vector>
#include <random>
#include <chrono>
#include <atlsafe.h>


// Import the Origin Automation Server type library and assign it
// to the origin namespace.
#import "Origin8.tlb" no_dual_interfaces rename_namespace("origin")


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
}


int main()
{
	// We'll use the Origin Automation Server for this example.

	HRESULT hr = ::CoInitializeEx(nullptr, COINIT_DISABLE_OLE1DDE | COINIT_MULTITHREADED);
	CLSID clsid;
	hr = ::CLSIDFromProgID(L"Origin.Application", &clsid);
	CComPtr<origin::IOApplication> app;
	hr = ::CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&app);
	app->Visible = origin::MAINWND_SHOW;

	origin::WorksheetPagePtr wksp = app->WorksheetPages->Add();
	origin::WorksheetPtr wks = wksp->Layers->Item[0];
	wks->Cols = 5;
	origin::ColumnPtr col_1 = wks->Columns->Item[0];
	origin::ColumnPtr col_2 = wks->Columns->Item[1];
	origin::ColumnPtr col_3 = wks->Columns->Item[2];
	origin::ColumnPtr col_4 = wks->Columns->Item[3];
	origin::ColumnPtr col_5 = wks->Columns->Item[4];


	// Begin examples.

	/*
	Mapping of supported C++ data type to supported Origin column data types and vice versa.
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

	// Vector and array functions:
	// void set_column_data(const ColumnPtr& ptr, const std::vector<T>& data, const std::size_t& offset = 0)
	// void set_column_data(const ColumnPtr& ptr, const T* data, const std::size_t& rows, const std::size_t& offset = 0)
	// Array version only for numeric and complex- not string data.

	// When setting column data, the column's data type will be changed to match
	// the first one that maps to the C++ data type. This can be turned off by:
	// #define ORGLAB_DATA_NO_CHANGE_DATA_TYPE

	// Exception thrown if ColumnPtr is invalid or can't set data for some reason.

	std::vector<double> vec_1 = { 1,2,2.3,3.4,4.5,5.6 };
	orglab_data::set_column_data(col_1, vec_1);

	std::vector<unsigned long> vec_2 = { 1, 2, 3, 4, 5 };
	orglab_data::set_column_data(col_2, vec_2);

	// std::wstring.
	std::vector<std::wstring> vec_3 = { L"hello world", L"مرحبا بالعالم", L"Բարեւ աշխարհ",
		L"Здравей свят", L"Прывітанне Сусвет", L"မင်္ဂလာပါကမ္ဘာလောက", L"你好，世界",
		L"Γειά σου Κόσμε", L"હેલ્લો વિશ્વ", L"Helló Világ", L"こんにちは世界", L"안녕 세상",
		L"سلام دنیا", L"העלא וועלט" };
	orglab_data::set_column_data(col_3, vec_3);

	// std::string.
	// Below observe data starts at the 14th (zero-based) row in column (offset param = 14).
	// Much slower than wide strings because it must convert from string to wide string.
	std::vector<std::string> vec_4 = {"Simple string", "Another simple string"};
	orglab_data::set_column_data(col_3, vec_4, 14);

	// Array version.
	unsigned short arr[5] = { 123, 234, 345, 456, 567 };
	orglab_data::set_column_data(col_4, arr, 5);

	std::vector<std::complex<double>> vec_5 = { std::complex<double>(1, 2),
		std::complex<double>(3, 4), std::complex<double>(5, 6) };
	orglab_data::set_column_data(col_5, vec_5);


	// Getting column data.

	// Functions:
	// std::vector<T> get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1)

	// When getting data from a column, you must assign it to a vector whose
	// C++ data type is mapped to the relevant column data type mapped above.
	// Otherwise, an exception is thrown.
	// Observe how type hints are used.

	// Exception thrown if ColumnPtr is invalid or an incompatible data type is specified.

	std::vector<double> vec_6 = orglab_data::get_column_data<double>(col_1);
	std::vector<unsigned long> vec_7 = orglab_data::get_column_data<unsigned long>(col_2);

	// Below observe offset & number of rows to return.
	std::vector<std::wstring> vec_8 = orglab_data::get_column_data<std::wstring>(col_3, 2, 3);
	std::vector<std::string> vec_9 = orglab_data::get_column_data<std::string>(col_3, 4, 5);

	// In this case, an exception will be throw because the specified data type does not match the column data type.
	try {
		std::vector<int> vec_10 = orglab_data::get_column_data<int>(col_1);
	}
	catch (const std::exception& e) {
		std::string s = e.what();
	}


	// How about a performance test?!?
	wks->Cols = wks->Cols++;
	origin::ColumnPtr col_1e6 = wks->Columns->Item[wks->Cols-1];
	std::vector<double> vec_in_1e6 = my_utils::get_test_data<double>(1e6);

	my_utils::elapsed_ms(true);
	orglab_data::set_column_data<double>(col_1e6, vec_in_1e6);
	std::cout << "Write 1E6 rows: " << my_utils::elapsed_ms().count() << " ms" << std::endl;
	my_utils::elapsed_ms(true);
	std::vector<double> vec_out_1e6 = orglab_data::get_column_data<double>(col_1e6);
	std::cout << "Read 1E6 rows: " << my_utils::elapsed_ms().count() << " ms" << std::endl;

	// End examples.


	std::cout << "Press any key to close this app.\n";
	char ch = _getch();

	app->Exit();
	app.Release();
	app = nullptr;
	::Sleep(500);
	::CoUninitialize();

}
