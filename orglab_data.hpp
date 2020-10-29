/*
MIT License

Copyright(c) 2020 OriginLab Corporation

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files(the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and /or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions :

The above copyright noticeand this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
#ifndef ORGLAB_DATA_HPP
#define ORGLAB_DATA_HPP

#include <exception>
#include <vector>
#include <complex>
#include <windows.h>
#include <comutil.h>
#include <comdef.h>
#include <atlsafe.h>
#include <ostream>
#include <assert.h>

#ifndef ORGLAB_DATA_ORGLAB_NS
#define ORGLAB_DATA_ORGLAB_NS origin
#endif

#ifdef ORGLAB_DATA_NO_CHANGE_DATA_TYPE
#define ORGLAB_DATA_CDT false
#else
#define ORGLAB_DATA_CDT true
#endif

namespace orglab_data {
	using namespace ORGLAB_DATA_ORGLAB_NS;

	//// PUBLIC API AT END OF FILE ////

	namespace impl { // Begin namespace for internal implementation.

		/* Simple 2D C++ "matrix" class */
		/* Meant to make it easier to interact with MatrixObjectPtr */
		template<class T>
		class matrix_adapter {
		protected:
			unsigned short rows_, cols_;
			T fill_;
			std::vector<T> vec_;
		public:
			/* Constructor */
			explicit matrix_adapter(const T& fill = T()) : rows_(0), cols_(0), fill_(fill) {}

			/* Constructor */
			explicit matrix_adapter(const unsigned short& rows, const unsigned short& cols, const T& fill = T()) :
				rows_(rows), cols_(cols), fill_(fill) {
				long sz = (long)(rows_ * cols_); // Avoids possible arith overflow per VS code analysis.
				vec_.resize(sz);
				std::fill(vec_.begin(), vec_.end(), fill_);
			}

			/* Constructor assigns raw array copying data */
			explicit matrix_adapter(T* data, const unsigned short& rows, const unsigned short& cols)
				: rows_(rows), cols_(cols) {
				long sz = (long)(rows_ * cols_);
				vec_.assign(data, data + sz);
			}

			~matrix_adapter() {}

			matrix_adapter(const matrix_adapter& other) : rows_(other.rows_),
				cols_(other.cols_), fill_(other.fill_), vec_(other.vec_) {}
			matrix_adapter(matrix_adapter&& other) noexcept : rows_(std::exchange(other.rows_, 0)),
				cols_(std::exchange(other.cols_, 0)), fill_(std::exchange(other.fill_, T())),
				vec_(std::move(other.vec_)) {}
			matrix_adapter& operator=(const matrix_adapter& other) noexcept {
				if (this != &other) {
					rows_ = other.rows_;
					cols_ = other.cols_;
					fill_ = other.fill_;
					vec_ = other.vec_;
				}
				return *this;
			}
			matrix_adapter& operator=(matrix_adapter&& other) noexcept {
				if (this != &other) {
					rows_ = std::exchange(other.rows_, 0);
					cols_ = std::exchange(other.cols_, 0);
					fill_ = std::exchange(other.fill_, T());
					vec_ = std::move(other.vec_);
				}
				return *this;
			}

			/* Assignment operator. Returns reference for given row and column */
			inline T& operator() (const unsigned short& row, const unsigned short& col) {
				assert(row < rows_ && col < cols_);
				long idx = (long)(rows_ * col + row); // Avoids possible arith overflow per VS code analysis.
				return vec_[idx];
			}

			/* Read operator. Returns value for given row and column */
			inline T operator() (const unsigned short& row, const unsigned short& col) const {
				assert(row < rows_ && col < cols_);
				long idx = (long)(rows_ * col + row); // Avoids possible arith overflow per VS code analysis.
				return vec_[idx];
			}

			/* Assigns raw array to matrix adapter */
			matrix_adapter& assign(T* data, const unsigned short& rows, const unsigned short& cols) {
				rows_ = rows;
				cols_ = cols;
				long sz = (long)(rows_ * cols_); // Avoids possible arith overflow per VS code analysis.
				vec_.assign(data, data + sz);
				vec_.shrink_to_fit();
				return *this;
			}

			/* Returns raw const array of internal storage */
			const T* data() const {
				return vec_.data();
			}

			/* Returns raw non-const (editable) array of internal storage */
			/* Useful when third-party library like Armadillo or Eigen assumes */
			/* control of the internal storage. */
			// Must not modify size of returned array */
			T* data() {
				return vec_.data();
			}

			/* Returns number of rows */
			inline unsigned short rows() const {
				return rows_;
			}

			/* Returns number of columns */
			inline unsigned short cols() const {
				return cols_;
			}

			/* Sets number of rows */
			/* Useful only when third-party library like Armadillo or Eigen assumes */
			/* control of the internal storage and has changed the shape of */
			/* storage. */
			matrix_adapter& rows(const unsigned short& rows) {
				if (0 == rows) {
					cols_ = rows_ * cols_;
					rows_ = 0;
				}
				else if (rows > rows_ * cols_) {
					rows_ = rows_ * cols_;
					cols_ == 0;
				}
				else {
					rows_ = rows;
					cols_ = (rows_ * cols_) / rows_;
				}
				return *this;
			}

			/* Sets number of columns */
			/* Useful only when third-party library like Armadillo or Eigen assumes */
			/* control of the internal storage and has changed the shape of */
			/* storage. */
			matrix_adapter& cols(const unsigned short& cols) {
				if (0 == cols) {
					rows_ = rows_ * cols_;
					cols_ = 0;
				}
				else if (cols > rows_ * cols_) {
					cols_ = rows_ * cols_;
					rows_ == 0;
				}
				else {
					cols_ = cols;
					rows_ = (rows_ * cols_) / cols_;
				}
				return *this;
			}

			/* Return size in elements of internal storage */
			inline std::size_t size() const {
				return vec_.size();
			}

			/* Changes dimensions of matrix adapter clearing data */
			matrix_adapter& resize(const unsigned short& rows, const unsigned short& cols) {
				rows_ = rows;
				cols_ = cols;
				long sz = (long)(rows * cols); // Avoids possible arith overflow per VS code analysis.
				std::vector<T>(sz).swap(vec_);
				std::fill(vec_.begin(), vec_.end(), fill_);
				return *this;
			}

			/* Resets matrix adapter */
			matrix_adapter& clear() {
				rows_ = 0;
				cols_ = 0;
				std::vector<T>().swap(vec_);
				vec_.shrink_to_fit();
				return *this;
			}

			/* Transposes and returns copy of matrix adapter */
			template <class T>
			matrix_adapter transpose() const {
				if (0 == vec_.size())
					return matrix_adapter<T>();
				// Clockwise rotation + horizontal flip.
				std::vector<T> v(vec_.size());
				long idx1 = 0, idx2 = 0;
				for (unsigned short row = 0; row < rows_; ++row) {
					for (unsigned short col = 0; col < cols_; ++col) {
						idx1 = (long)(cols_ * row + col);
						idx2 = (long)(rows_ * col + row);
						v[idx1] = vec_[idx2];
					}
				}
				return matrix_adapter<T>(v.data(), cols_, rows_); // Reverse.
			}

			/* Transposes matrix adapter in place */
			matrix_adapter& transpose_self() {
				if (0 == vec_.size())
					return *this;
				// Clockwise rotation + horizontal flip.
				std::vector<T> v(vec_);
				long idx1 = 0, idx2 = 0;
				for (unsigned short row = 0; row < rows_; ++row) {
					for (unsigned short col = 0; col < cols_; ++col) {
						idx1 = (long)(cols_ * row + col);
						idx2 = (long)(rows_ * col + row);
						vec_[idx1] = v[idx2];
					}
				}
				unsigned short r = rows_;
				rows_ = cols_;
				cols_ = r;
				return *this;
			}

			/* Iterator */
			using iterator = typename std::vector<T>::iterator;
			/* Const iterator */
			using const_iterator = typename std::vector<T>::const_iterator;
			/* Iterator method */
			iterator begin() noexcept { return vec_.begin(); }
			/* Iterator method */
			iterator end() noexcept { return vec_.end(); }
			/* Iterator method */
			const_iterator cbegin() const noexcept { return vec_.cbegin(); }
			/* Iterator method */
			const_iterator cend() const noexcept { return vec_.cend(); }

		};

		/* Operator << for matrix_adapter */
		/* Dumps to an output stream the contents of the matrix_adapter object */
		/* E.g. std::cout << "Matrix:\n" << ma; */
		template<class T, typename Char, typename Traits>
		std::basic_ostream<typename Char, typename Traits>& operator<< (std::basic_ostream<typename Char, typename Traits>& out, const matrix_adapter<T>& ma)
		{
			unsigned short rows = ma.rows(), cols = ma.cols();
			if (rows * cols < 1)
				return out;
			for (unsigned short i = 0; i < rows; ++i) {
				for (unsigned short j = 0; j < cols; ++j) {
					if ((cols - 1) == j)
						out << ma(i, j);
					else
						out << ma(i, j) << "\t";
				}
				if ((rows - 1) != i)
					out << "\n";
			}
			return out;
		}

		inline std::string from_wide(const std::wstring& wstr) {
			if (wstr.empty())
				return std::string();
			int len = ::WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), static_cast<int>(wstr.size()), NULL, 0, NULL, NULL);
			if (len < 1)
				return std::string();
			std::string str(len, '\0');
			if (0 == ::WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), static_cast<int>(wstr.size()), &str[0], static_cast<int>(str.size()), NULL, NULL))
				return std::string();
			return str;
		}

		inline std::wstring to_wide(const std::string& str) {
			if (str.empty())
				return std::wstring();
			int len = ::MultiByteToWideChar(CP_UTF8, 0, str.c_str(), static_cast<int>(str.size()), NULL, 0);
			if (len < 1)
				return std::wstring();
			std::wstring wstr(len, '\0');
			if (0 == ::MultiByteToWideChar(CP_UTF8, 0, str.c_str(), static_cast<int>(str.size()), &wstr[0], static_cast<int>(wstr.size())))
				return std::wstring();
			return wstr;
		}

		inline std::wstring from_bstr(const BSTR& bstr) {
			return std::wstring(bstr, ::SysStringLen(bstr));
		}

		inline CComBSTR to_ccom_bstr(const std::wstring& wstr) {
			if (wstr.empty())
				return CComBSTR();
			return CComBSTR(static_cast<int>(wstr.size()), wstr.data());
		}

		inline CComBSTR to_ccom_bstr(const std::string& str) {
			return to_ccom_bstr(to_wide(str));
		}

		template <class T, typename std::enable_if<std::is_integral<T>::value>::type * = nullptr>
		inline long to_non_negative_long(const T& t) {
			if (t > LONG_MAX)
				return LONG_MAX;
			if (t < 0)
				return 0;
			return static_cast<long>(t);
		}

		template <class T>
		inline unsigned short to_unsigned_short(const T& t) {
			static const T max_val = static_cast<T>((std::numeric_limits<unsigned short>::max)());
			static const T min_val = static_cast<T>((std::numeric_limits<unsigned short>::min)());
			if (t < min_val)
				return (std::numeric_limits<unsigned short>::min)();
			if (t > max_val)
				return (std::numeric_limits<unsigned short>::max)();
			return static_cast<unsigned short>(t);
		}

		using com_compat_info_t = std::pair<COLDATAFORMAT, VARENUM>;

		template< class T>
		inline com_compat_info_t get_com_compat_info(const COLDATAFORMAT& fmt, bool is_matrix = false) {
			if (COLDATAFORMAT::DF_DATE == fmt || COLDATAFORMAT::DF_TIME == fmt) return com_compat_info_t{ fmt, VT_R8 };
			if (std::is_same<T, double>::value)						return com_compat_info_t{ is_matrix ? COLDATAFORMAT::DF_DOUBLE : COLDATAFORMAT::DF_TEXT_NUMERIC, VT_R8 };
			if (std::is_same<T, float>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_FLOAT, VT_R4 };
			if (std::is_same<T, int>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_LONG, VT_I4 };
			if (std::is_same<T, long>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_LONG, VT_I4 };
			if (std::is_same<T, unsigned long>::value)				return com_compat_info_t{ COLDATAFORMAT::DF_ULONG, VT_I4 };
			if (std::is_same<T, short>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_SHORT, VT_I2 };
			if (std::is_same<T, unsigned short>::value)				return com_compat_info_t{ COLDATAFORMAT::DF_USHORT, VT_I2 };
			if (std::is_same<T, std::wstring>::value)				return com_compat_info_t{ COLDATAFORMAT::DF_TEXT_NUMERIC, VT_BSTR };
			if (std::is_same<T, std::string>::value)				return com_compat_info_t{ COLDATAFORMAT::DF_TEXT_NUMERIC, VT_BSTR };
			if (std::is_same<T, byte>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_BYTE, VT_I1 };
			if (std::is_same<T, char>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_CHAR, VT_I1 };
			if (std::is_same<T, std::complex<double>>::value)		return com_compat_info_t{ COLDATAFORMAT::DF_COMPLEX, VT_R8 };

			throw std::exception("Incompatible data types");
		}

		template<class T>
		inline bool is_vector_type_compatible(const COLDATAFORMAT& fmt) {
			if (std::is_same<T, double>::value&& COLDATAFORMAT::DF_TEXT_NUMERIC == fmt)			return true;
			if (std::is_same<T, double>::value&& COLDATAFORMAT::DF_DOUBLE == fmt)				return true;
			if (std::is_same<T, double>::value&& COLDATAFORMAT::DF_DATE == fmt)					return true;
			if (std::is_same<T, double>::value&& COLDATAFORMAT::DF_TIME == fmt)					return true;
			if (std::is_same<T, float>::value&& COLDATAFORMAT::DF_FLOAT == fmt)					return true;
			if (std::is_same<T, int>::value&& COLDATAFORMAT::DF_LONG == fmt)					return true;
			if (std::is_same<T, long>::value&& COLDATAFORMAT::DF_LONG == fmt)					return true;
			if (std::is_same<T, unsigned long>::value&& COLDATAFORMAT::DF_ULONG == fmt)			return true;
			if (std::is_same<T, short>::value&& COLDATAFORMAT::DF_SHORT == fmt)					return true;
			if (std::is_same<T, unsigned short>::value&& COLDATAFORMAT::DF_USHORT == fmt)		return true;
			if (std::is_same<T, std::wstring>::value&& COLDATAFORMAT::DF_TEXT_NUMERIC == fmt)	return true;
			if (std::is_same<T, std::string>::value&& COLDATAFORMAT::DF_TEXT_NUMERIC == fmt)	return true;
			if (std::is_same<T, std::wstring>::value&& COLDATAFORMAT::DF_TEXT == fmt)			return true;
			if (std::is_same<T, std::string>::value&& COLDATAFORMAT::DF_TEXT == fmt)			return true;
			if (std::is_same<T, byte>::value&& COLDATAFORMAT::DF_BYTE == fmt)					return true;
			if (std::is_same<T, char>::value&& COLDATAFORMAT::DF_CHAR == fmt)					return true;
			if (std::is_same<T, std::complex<double>>::value&& COLDATAFORMAT::DF_COMPLEX == fmt)	return true;
			return false;
		}

		void do_set_col_data(const ColumnPtr& col, const _variant_t& vt_array, const long& offset) {
			try {
				_variant_t v_offset(offset);
				col->SetData(vt_array, v_offset);
			}
			catch (...) {
				throw std::exception("ColumnPtr set data fail");
			}
		}

		_variant_t do_get_col_data(const ColumnPtr& col, const ARRAYDATAFORMAT& fmt, const long& offset, const long& rows) {
			if (0 == rows)
				return _variant_t();
			long r2 = rows < -1 ? -1 : rows;
			if (r2 > 0)
				r2 = offset + r2 - 1;
			_variant_t v_r1(offset);
			_variant_t v_r2(r2);
			_variant_t v_lbound(0);
			return col->GetData(fmt, v_r1, v_r2, v_lbound);
		}

		template<class T>
		void set_arithmetic_column_data(const ColumnPtr& col, const T* data, const std::size_t& rows, const std::size_t& offset, bool change_type = true) {
			if (!data || 0 == rows)
				return;
			com_compat_info_t info = get_com_compat_info<T>(col->DataFormat);
			if (change_type && (info.first != col->DataFormat))
				col->DataFormat = info.first;
			try {
				long long_rows = to_non_negative_long(rows);
				SAFEARRAYBOUND sa_bounds = { static_cast<unsigned long>(long_rows), 0 };
				SAFEARRAY* pSA = ::SafeArrayCreate(info.second, 1, &sa_bounds);
				_variant_t vt_array;
				vt_array.vt = info.second | VT_ARRAY;
				vt_array.parray = pSA; // Let _variant_t take ownership of SafeArray.
				T* p_val = nullptr;
				::SafeArrayAccessData(pSA, (void**)&p_val);
				memcpy(p_val, data, long_rows * sizeof(T));
				::SafeArrayUnaccessData(pSA);
				do_set_col_data(col, vt_array, to_non_negative_long(offset));
			}
			catch (...) {
				throw std::exception("ColumnPtr set data fail");
			}
		}

		void set_complex_column_data(const ColumnPtr& col, const std::complex<double>* data, const std::size_t& rows, const std::size_t& offset, bool change_type = true) {
			if (!data || 0 == rows)
				return;
			com_compat_info_t info = get_com_compat_info<std::complex<double>>(col->DataFormat);
			if (change_type && (info.first != col->DataFormat))
				col->DataFormat = info.first;
			try {
				long long_rows = to_non_negative_long(rows);
				SAFEARRAYBOUND sa_bounds = { static_cast<unsigned long>(long_rows) * 2, 0 };
				SAFEARRAY* pSA = ::SafeArrayCreate(info.second, 1, &sa_bounds);
				_variant_t vt_array;
				vt_array.vt = info.second | VT_ARRAY;
				vt_array.parray = pSA; // Let _variant_t take ownership of SafeArray.
				double* p_val = nullptr;
				::SafeArrayAccessData(pSA, (void**)&p_val);
				for (std::size_t i = 0; i < long_rows; ++i) {
					*p_val = data[i].real();
					*(++p_val) = data[i].imag();
					++p_val;
				}
				::SafeArrayUnaccessData(pSA);
				do_set_col_data(col, vt_array, to_non_negative_long(offset));
			}
			catch (...) {
				throw std::exception("ColumnPtr set data fail");
			}
		}

		template<class T>
		void set_string_column_data(const ColumnPtr& col, const std::vector<T>& data, const std::size_t& offset, bool change_type = true) {
			if (0 == data.size())
				return;
			com_compat_info_t info = get_com_compat_info<T>(col->DataFormat);
			if (change_type && (info.first != col->DataFormat))
				col->DataFormat = info.first;
			try {
				long long_rows = to_non_negative_long(data.size());
				CComSafeArray<BSTR> csa(long_rows);
				for (long i = 0; i < long_rows; i++) {
					csa.SetAt(i, to_ccom_bstr(data[i]).Detach(), false);
				}
				_variant_t vt_array;
				vt_array.vt = VT_BSTR | VT_ARRAY;
				vt_array.parray = csa.Detach(); // Let _variant_t take ownership of CComSafeArray's SAFEARRAY.

				do_set_col_data(col, vt_array, to_non_negative_long(offset));
			}
			catch (...) {
				throw std::exception("ColumnPtr set data fail");
			}
		}

		template<class T>
		void get_arithmetic_column_data(const ColumnPtr& col, std::vector<T>& data, const long& offset, const long& rows) {
			if (!is_vector_type_compatible<T>(col->DataFormat))
				throw std::exception("Incompatible data types");
			_variant_t vt_data = do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_NUMERIC, to_non_negative_long(offset), rows < -1 ? -1 : rows);
			if (VT_ARRAY & vt_data.vt) {
				long lbound, ubound;
				::SafeArrayGetLBound(vt_data.parray, 1, &lbound);
				::SafeArrayGetUBound(vt_data.parray, 1, &ubound);
				long count = ubound - lbound + 1;
				if (count > 0) {
					data.reserve(count);
					T* p_val = nullptr;
					::SafeArrayAccessData(vt_data.parray, (void**)&p_val);
					data.assign(p_val, p_val + count);
					::SafeArrayUnaccessData(vt_data.parray);
				}
			}
		}

		void get_complex_column_data(const ColumnPtr& col, std::vector<std::complex<double>>& data, const long& offset, const long& rows) {
			_variant_t vt_data = do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_NUMERIC, to_non_negative_long(offset), rows < -1 ? -1 : rows);
			if (VT_ARRAY & vt_data.vt) {
				long lbound, ubound;
				::SafeArrayGetLBound(vt_data.parray, 1, &lbound);
				::SafeArrayGetUBound(vt_data.parray, 1, &ubound);
				long count = ubound - lbound + 1;
				if (count > 0) {
					// OrgLab returns pattern below for complex data:
					// c1.re,c1.im,c2.re,c2.im,c3.re,c3.im, etc.
					// There will be twice as many values as is needed
					// for a complex vector since complex has two parts.
					// Use resize to already add default complex values
					// to vector.
					data.resize(count / 2);
					double* p_val = nullptr;
					::SafeArrayAccessData(vt_data.parray, (void**)&p_val);
					for (std::size_t i = 0; i < data.size(); ++i, ++p_val) {
						data[i].real(*p_val);
						data[i].imag(*(++p_val));
					}
					::SafeArrayUnaccessData(vt_data.parray);
				}
			}
		}

		void get_wstring_column_data(const ColumnPtr& col, std::vector<std::wstring>& data, const long& offset, const long& rows) {
			if (!is_vector_type_compatible<std::wstring>(col->DataFormat))
				throw std::exception("Incompatible data types");
			CComSafeArray<BSTR> csa;
			{ // This scope makes sure vt_data is cleaned up quickly for performance.
				_variant_t vt_data = do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_STR, to_non_negative_long(offset), rows < -1 ? -1 : rows);
				if (VT_ARRAY & vt_data.vt) {
					csa.Attach(vt_data.parray);
					vt_data.Detach();
				}
				else
					return;
			}
			std::size_t count = csa.GetCount(0);
			if (0 == count)
				return;
			data.resize(count);
			::SafeArrayLock(csa.m_psa);
			BSTR* p_csa = (BSTR*)(csa.m_psa->pvData);
			for (long i = 0; i < count; ++i, p_csa++) {
				data[i] = from_bstr(*p_csa);
			}
			::SafeArrayUnlock(csa.m_psa);
		}

		void get_string_column_data(const ColumnPtr& col, std::vector<std::string>& data, const long& offset, const long& rows) {
			if (!is_vector_type_compatible<std::wstring>(col->DataFormat))
				throw std::exception("Incompatible data types");
			CComSafeArray<BSTR> csa;
			{ // This scope makes sure vt_data is cleaned up quickly for performance.
				_variant_t vt_data = do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_STR, to_non_negative_long(offset), rows < -1 ? -1 : rows);
				if (VT_ARRAY & vt_data.vt) {
					csa.Attach(vt_data.parray);
					vt_data.Detach();
				}
				else
					return;
			}
			std::size_t count = csa.GetCount(0);
			if (0 == count)
				return;
			data.resize(count);
			::SafeArrayLock(csa.m_psa);
			BSTR* p_csa = (BSTR*)(csa.m_psa->pvData);
			for (long i = 0; i < count; ++i, p_csa++) {
				data[i] = from_wide(from_bstr(*p_csa));
			}
			::SafeArrayUnlock(csa.m_psa);
		}

		void do_set_mat_data(const MatrixObjectPtr& mat, const _variant_t& vt_array) {
			try {
				_variant_t v_zero(0);
				mat->SetData(vt_array, v_zero, v_zero);
			}
			catch (...) {
				throw std::exception("MatrixObjectPtr set data fail");
			}
		}

		_variant_t do_get_mat_data(const MatrixObjectPtr& mat, const ARRAYDATAFORMAT& fmt) {
			_variant_t v_r1(0);
			_variant_t v_r2(-1);
			_variant_t v_c1(0);
			_variant_t v_c2(-1);
			_variant_t v_lbound(0);
			return mat->GetData(v_r1, v_c1, v_r2, v_c2, fmt, v_lbound);
		}

		template<class T>
		void set_arithmetic_matrix_data(const MatrixObjectPtr& mat, const matrix_adapter<T> ma, bool change_type = true) {
			unsigned short rows = ma.rows();
			unsigned short cols = ma.cols();
			const T* data = ma.data();
			if (!data || 0 == rows * cols)
				return;
			com_compat_info_t info = get_com_compat_info<T>(mat->DataFormat, true);
			if (change_type && (info.first != mat->DataFormat))
				mat->DataFormat = info.first;
			try {
				SAFEARRAYBOUND sa_bounds[2];
				sa_bounds[0].lLbound = 0;
				sa_bounds[0].cElements = cols; //rows;
				sa_bounds[1].lLbound = 0;
				sa_bounds[1].cElements = rows; //cols;
				SAFEARRAY* pSA = ::SafeArrayCreate(info.second, 2, sa_bounds);
				_variant_t vt_array;
				vt_array.vt = info.second | VT_ARRAY;
				vt_array.parray = pSA; // Let _variant_t take ownership of SafeArray.

				T* p_vals = nullptr;
				::SafeArrayAccessData(pSA, (void**)&p_vals);
				// In recent versions of Origin, MatrixObjectPtr became column major. Hence need to transpose.
				//memcpy(p_vals, data, static_cast<std::size_t>(rows)* static_cast<std::size_t>(cols) * sizeof(T));
				for (unsigned short k = 0; k < cols; ++k) {
					for (unsigned short j = 0; j < rows; ++j) {
						p_vals[cols * j + k] = data[rows * k + j];
					}
				}
				::SafeArrayUnaccessData(pSA);
				do_set_mat_data(mat, vt_array);
			}
			catch (...) {
				throw std::exception("MatrixObjectPtr set data fail");
			}
		}

		void set_complex_matrix_data(const MatrixObjectPtr& mat, const matrix_adapter<std::complex<double>> ma, bool change_type = true) {
			unsigned short rows = ma.rows();
			unsigned short cols = ma.cols();
			const std::complex<double>* data = ma.data();
			if (!data || 0 == rows * cols)
				return;
			com_compat_info_t info = get_com_compat_info<std::complex<double>>(mat->DataFormat, true);
			if (change_type && (info.first != mat->DataFormat))
				mat->DataFormat = info.first;
			try {
				SAFEARRAYBOUND sa_bounds[3];
				sa_bounds[0].lLbound = 0;
				sa_bounds[0].cElements = cols; //rows;
				sa_bounds[1].lLbound = 0;
				sa_bounds[1].cElements = rows; //cols;
				sa_bounds[2].lLbound = 0;
				sa_bounds[2].cElements = 2;
				SAFEARRAY* pSA = ::SafeArrayCreate(info.second, 3, sa_bounds);
				_variant_t vt_array;
				vt_array.vt = info.second | VT_ARRAY;
				vt_array.parray = pSA; // Let _variant_t take ownership of SafeArray.
				double* p_vals = nullptr;
				::SafeArrayAccessData(pSA, (void**)&p_vals);
				// In recent versions of Origin, MatrixObjectPtr became column major. Hence need to transpose.
				/*
				for (unsigned short k = 0; k < cols; ++k) {
					for (unsigned short j = 0; j < rows; ++j, ++p_vals) {
						*p_vals = data[j,k].real();
					}
				}
				for (unsigned short k = 0; k < cols; ++k) {
					for (unsigned short j = 0; j < rows; ++j, ++p_vals) {
						*p_vals = data[j,k].imag();
					}
				}
				*/
				for (unsigned short k = 0; k < cols; ++k) {
					for (unsigned short j = 0; j < rows; ++j) {
						p_vals[cols * j + k] = data[rows * k + j].real();
					}
				}
				unsigned short i = cols * rows;
				for (unsigned short k = 0; k < cols; ++k) {
					for (unsigned short j = 0; j < rows; ++j) {
						p_vals[cols * j + k + i] = data[rows * k + j].imag();
					}
				}

				::SafeArrayUnaccessData(pSA);
				do_set_mat_data(mat, vt_array);
			}
			catch (...) {
				throw std::exception("MatrixObjectPtr set data fail");
			}
		}

		template<class T>
		void get_arithmetic_matrix_data(const MatrixObjectPtr& mat, matrix_adapter<T>& ma) {
			if (!is_vector_type_compatible<T>(mat->DataFormat))
				throw std::exception("Incompatible data types");
			_variant_t vt_data = do_get_mat_data(mat, ARRAYDATAFORMAT::ARRAY2D_NUMERIC);
			if (VT_ARRAY & vt_data.vt) {
				long lbound1, ubound1, lbound2, ubound2;
				::SafeArrayGetLBound(vt_data.parray, 1, &lbound1);
				::SafeArrayGetUBound(vt_data.parray, 1, &ubound1);
				::SafeArrayGetLBound(vt_data.parray, 2, &lbound2);
				::SafeArrayGetUBound(vt_data.parray, 2, &ubound2);
				long count1 = ubound1 - lbound1 + 1;
				long count2 = ubound2 - lbound2 + 1;
				if (count1 > 0 && count2 > 0) {
					T* p_val = nullptr;
					::SafeArrayAccessData(vt_data.parray, (void**)&p_val);
					ma.assign(p_val, to_unsigned_short(count1), to_unsigned_short(count2));
					::SafeArrayUnaccessData(vt_data.parray);
				}
			}
		}

		void get_complex_matrix_data(const MatrixObjectPtr& mat, matrix_adapter<std::complex<double>>& ma) {
			CComSafeArray<double> csa;
			{ // Scope releases vt_data as soon as we hand its SAFEARRAY off to the CComSafeArray.
				_variant_t vt_data = do_get_mat_data(mat, ARRAYDATAFORMAT::ARRAY2D_NUMERIC);
				if (VT_ARRAY & vt_data.vt) {
					csa.Attach(vt_data.parray);
					vt_data.Detach();
				}
			}
			// Despite ARRAY2D_NUMERIC, 3D array is returned for complex.
			// CComSafeArray handles such arrays easier, that's why we use one. Little to no overhead.
			unsigned short rows = to_unsigned_short(csa.GetCount(0));
			unsigned short cols = to_unsigned_short(csa.GetCount(1));
			unsigned short parts = to_unsigned_short(csa.GetCount(2)); // Two parts- 1st is real part, 2nd is imaginary part.
			// There can be no empty matrix object in Origin- always have at least 1x1.
			if (rows > 0 && cols > 0 && parts > 0) {
				ma.resize(rows, cols);
				::SafeArrayLock(csa.m_psa);
				double* p_csa_real = (double*)(csa.m_psa->pvData); // Real part is first rows*cols values.
				long offset = (long)(rows * cols); // Avoids possible arith overflow per VS code analysis.
				double* p_csa_imag = p_csa_real + offset;// (rows * cols); // Imaginary part real part offset by rows*cols.
				for (unsigned short c = 0; c < cols; ++c) {
					for (unsigned short r = 0; r < rows; ++r, p_csa_real++, p_csa_imag++) {
						ma(r, c) = std::complex<double>{ *p_csa_real , *p_csa_imag };
					}
				}
				::SafeArrayUnlock(csa.m_psa);
			}
		}

	} // End namespace for internal implementation.

	//// BEGIN PUBLIC API ////

	using impl::matrix_adapter;

	_bstr_t to_str_prop(const std::wstring& str) {
		return str.c_str();
	}
	_bstr_t to_str_prop(const std::string& str) { return to_str_prop(impl::to_wide(str)); }

	template<class T> T from_str_prop(const _bstr_t& prop);
	template<> std::wstring from_str_prop(const _bstr_t& prop) {
		return std::wstring{ prop, ::SysStringLen(prop) };
	}
	template<> std::string from_str_prop(const _bstr_t& prop) {
		return impl::from_wide(from_str_prop<std::wstring>(prop));
	}

	template<class T>
	typename std::enable_if<std::is_arithmetic<T>::value, void>::type
		set_column_data(const ColumnPtr& ptr, const std::vector<T>& data, const std::size_t& offset = 0) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		impl::set_arithmetic_column_data(ptr, data.data(), data.size(), offset, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_arithmetic<T>::value, void>::type
		set_column_data(const ColumnPtr& ptr, const T* data, const std::size_t& rows, const std::size_t& offset = 0) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		impl::set_arithmetic_column_data(ptr, data, rows, offset, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::complex<double>>::value, void>::type
		set_column_data(const ColumnPtr& ptr, const std::vector<T>& data, const std::size_t& offset = 0) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		impl::set_complex_column_data(ptr, data.data(), data.size(), offset, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::complex<double>>::value, void>::type
		set_column_data(const ColumnPtr& ptr, const T* data, const std::size_t& rows, const std::size_t& offset = 0) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		impl::set_complex_column_data(ptr, data, rows, offset, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::wstring>::value || std::is_same<T, std::string>::value, void>::type
		set_column_data(const ColumnPtr& ptr, const std::vector<T>& data, const std::size_t& offset = 0) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		impl::set_string_column_data<T>(ptr, data, offset, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_arithmetic<T>::value, std::vector<T>>::type
		get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		std::vector<T> data;
		impl::get_arithmetic_column_data<T>(ptr, data, offset, rows);
		return data;
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::complex<double>>::value, std::vector<T>>::type
		get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		std::vector<T> data;
		impl::get_complex_column_data(ptr, data, offset, rows);
		return data;
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::wstring>::value, std::vector<T>>::type
		get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		std::vector<T> data;
		impl::get_wstring_column_data(ptr, data, offset, rows);
		return data;
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::string>::value, std::vector<T>>::type
		get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		std::vector<T> data;
		impl::get_string_column_data(ptr, data, offset, rows);
		return data;
	}

	template<class T>
	typename std::enable_if<std::is_arithmetic<T>::value, void>::type
		set_matrix_data(const MatrixObjectPtr& ptr, const matrix_adapter<T>& ma) {
		if (!ptr)
			throw std::exception("MatrixObjectPtr is invalid");
		impl::set_arithmetic_matrix_data(ptr, ma, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::complex<double>>::value, void>::type
		set_matrix_data(const MatrixObjectPtr& ptr, const matrix_adapter<T>& ma) {
		if (!ptr)
			throw std::exception("MatrixObjectPtr is invalid");
		impl::set_complex_matrix_data(ptr, ma, ORGLAB_DATA_CDT);
	}

	template<class T>
	typename std::enable_if<std::is_arithmetic<T>::value, matrix_adapter<T>>::type
		get_matrix_data(const MatrixObjectPtr& ptr) {
		if (!ptr)
			throw std::exception("MatrixObjectPtr is invalid");
		matrix_adapter<T> ma;
		impl::get_arithmetic_matrix_data<T>(ptr, ma);
		return ma;
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::complex<double>>::value, matrix_adapter<T>>::type
		get_matrix_data(const MatrixObjectPtr& ptr) {
		if (!ptr)
			throw std::exception("MatrixObjectPtr is invalid");
		matrix_adapter<T> ma;
		impl::get_complex_matrix_data(ptr, ma);
		return ma;
	}

	//// END PUBLIC API ////

} /* End namespace orglab_data */

#endif /* ORGLAB_DATA_HPP */
