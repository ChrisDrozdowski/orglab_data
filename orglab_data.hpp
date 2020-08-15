#ifndef ORGLAB_DATA_HPP
#define ORGLAB_DATA_HPP

#include <vector>
#include <complex>
#include <windows.h>
#include <comutil.h>
#include <comdef.h>
#include <atlsafe.h>

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

	/* Begin forward declarations of implementation functions */
	namespace impl {
		template<class T> void set_arithmetic_column_data(const ColumnPtr& col, const T* data, const std::size_t& rows, const std::size_t& offset, bool change_type = true);
		void set_complex_column_data(const ColumnPtr& col, const std::complex<double>* data, const std::size_t& rows, const std::size_t& offset, bool change_type = true);
		template<class T> void set_string_column_data(const ColumnPtr& col, const std::vector<T>& data, const std::size_t& offset, bool change_type = true);
		template<class T> void get_arithmetic_column_data(const ColumnPtr& col, std::vector<T>& data, const long& offset, const long& rows);
		void get_complex_column_data(const ColumnPtr& col, std::vector<std::complex<double>>& data, const long& offset, const long& rows);
		void get_wstring_column_data(const ColumnPtr& col, std::vector<std::wstring>& data, const long& offset, const long& rows);
		void get_string_column_data(const ColumnPtr& col, std::vector<std::string>& data, const long& offset, const long& rows);
	}
	/* End forward declarations of implementation functions */


	/* Begin public API */

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
		std::vector<std::complex<double>> data;
		impl::get_complex_column_data(ptr, data, offset, rows);
		return data;
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::wstring>::value, std::vector<T>>::type
		get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		std::vector<std::wstring> data;
		impl::get_wstring_column_data(ptr, data, offset, rows);
		return data;
	}

	template<class T>
	typename std::enable_if<std::is_same<T, std::string>::value, std::vector<T>>::type
		get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1) {
		if (!ptr)
			throw std::exception("ColumnPtr is invalid");
		std::vector<std::string> data;
		impl::get_string_column_data(ptr, data, offset, rows);
		return data;
	}

	/* End public API */


	/* Begin implementation functions */
	namespace impl {

		// Supporting functions.

		template <class T, typename std::enable_if<std::is_integral<T>::value>::type * = nullptr>
		inline long to_non_negative_long(const T& t) {
			if (t > LONG_MAX)
				return LONG_MAX;
			if (t < 0)
				return 0;
			return static_cast<long>(t);
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

		// Main functionality.

		using com_compat_info_t = std::pair<COLDATAFORMAT, VARENUM>;

		template< class T>
		inline com_compat_info_t _get_com_compat_info(const COLDATAFORMAT& fmt) {
			if (COLDATAFORMAT::DF_DATE == fmt || COLDATAFORMAT::DF_TIME == fmt) return com_compat_info_t{ fmt, VT_R8 };
			if (std::is_same<T, double>::value)						return com_compat_info_t{ COLDATAFORMAT::DF_TEXT_NUMERIC, VT_R8 };
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

		void do_set_col_data(const ColumnPtr& col, const _variant_t& vt_array, const long& offset)
		{
			try {
				_variant_t v_offset(offset);
				col->SetData(vt_array, v_offset);
			}
			catch (...) {
				throw std::exception("ColumnPtr set data fail");
			}
		}

		template<class T>
		void set_arithmetic_column_data(const ColumnPtr& col, const T* data, const std::size_t& rows, const std::size_t& offset, bool change_type) {
			if (!data || 0 == rows)
				return;
			com_compat_info_t info = _get_com_compat_info<T>(col->DataFormat);
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

		void set_complex_column_data(const ColumnPtr& col, const std::complex<double>* data, const std::size_t& rows, const std::size_t& offset, bool change_type) {
			if (!data || 0 == rows)
				return;
			com_compat_info_t info = _get_com_compat_info<std::complex<double>>(col->DataFormat);
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
		void set_string_column_data(const ColumnPtr& col, const std::vector<T>& data, const std::size_t& offset, bool change_type) {
			if (0 == data.size())
				return;
			com_compat_info_t info = _get_com_compat_info<T>(col->DataFormat);
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
			if (std::is_same<T, std::complex<double>>::value&& COLDATAFORMAT::DF_COMPLEX == fmt)return true;
			return false;
		}

		_variant_t _do_get_col_data(const ColumnPtr& col, const ARRAYDATAFORMAT& fmt, const long& offset, const long& rows) {
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
		void get_arithmetic_column_data(const ColumnPtr& col, std::vector<T>& data, const long& offset, const long& rows) {
			if (!is_vector_type_compatible<T>(col->DataFormat))
				throw std::exception("Incompatible data types");
			_variant_t vt_data = _do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_NUMERIC, to_non_negative_long(offset), rows < -1 ? -1 : rows);
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
			_variant_t vt_data = _do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_NUMERIC, to_non_negative_long(offset), rows < -1 ? -1 : rows);
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
				_variant_t vt_data = _do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_STR, to_non_negative_long(offset), rows < -1 ? -1 : rows);
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
				_variant_t vt_data = _do_get_col_data(col, ARRAYDATAFORMAT::ARRAY1D_STR, to_non_negative_long(offset), rows < -1 ? -1 : rows);
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

	} /* End namespace impl */
	/* End implementation functions */

} /* End namespace orglab_data */

#endif /* ORGLAB_DATA_HPP */
