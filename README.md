# orglab_data

- Single file (header only library): [orglab_data.hpp](orglab_data.hpp)
- MSVC with C++17 or later required.
- 64-bit (x64) only.
- Can use either Origin Automation Server or Orglab.
- Tested with Origin 2021 installed whether Automation Server or Orglab used in code.
- Please, please look at the [example Visual Studio 2019 solution](orglab_data_example) for good examples and additional important information.

### API

#### Setup
Prior to using the API, you need to assign the `#ORGLAB_DATA_ORGLAB_NS` macro to the name of the namespace you used when importing the type library for either Origin Automation Server or OrgLab. You HAVE to do this. For example:

```cpp
#import "Origin8.tlb" no_dual_interfaces rename_namespace("origin")
// or
#import "orglab.tlb" no_dual_interfaces rename_namespace("origin")

// Now define in either case.
#define ORGLAB_DATA_ORGLAB_NS origin

#include "orglab_data.hpp"

```

#### Functions

```cpp
/* Sets (inserts) column data using a vector.
 *
 * Parameters
 *   ColumnPtr		ptr 	Instance representing a column in a worksheet.
 *   std::vector<T>	data	The data to insert. Supported C++ data types are listed below.
 *   std::size_t	offset	Zero-based row offset to start data insertion.
 *
 * Returns
 *   void
 *
 * Throws
 *   Throws std::exception if ColumnPtr instance is invalid or if data cannot be inserted
 *   for some unknown reason.
 *
 * Notes
 *   Supported C++ data types: double, float, int, long, unsigned long, short, unsigned short,
 *   std::wstring, std::string, byte, char, std::complex<double>.
 *
 *   Converts column type to the one mapped to the first C++ data type below. This behavior
 *   may be turned off by defining: ORGLAB_DATA_NO_CHANGE_DATA_TYPE
 */
template<class T> void set_column_data(const ColumnPtr& ptr, const std::vector<T>& data, const std::size_t& offset = 0)
```

```cpp
/* Sets (inserts) column data using a pointer to an array.
 *
 * Parameters
 *   ColumnPtr		ptr 	Instance representing a column in a worksheet.
 *   T*	data		data	Pointer to an array of data to insert. Supported C++ data types are listed below.
 *   std::size_t	rows	Size of array.
 *   std::size_t	offset	Zero-based row offset to start data insertion.
 *
 * Returns
 *   void
 *
 * Throws
 *   Throws std::exception if ColumnPtr instance is invalid or if data cannot be inserted
 *   for some unknown reason.
 *
 * Notes
 *   Supported C++ data types: double, float, int, long, unsigned long, short, unsigned short,
 *   byte, char, std::complex<double> (NOT std::wstring or std::string).
 *
 *   Converts column type to the first one mapped to the data type below. This behavior
 *   may be turned off by defining: ORGLAB_DATA_NO_CHANGE_DATA_TYPE
 */
template<class T> void set_column_data(const ColumnPtr& ptr, const T* data, const std::size_t& rows, const std::size_t& offset = 0)
```

```cpp
/* Gets (retrieves) a vector of data from a column.
 *
 * Parameters
 *   ColumnPtr	ptr 	Instance representing a column in a worksheet.
 *   long		offset	Zero-based row offset to start data retrieval.
 *   long		rows	Number of rows to return.
 *
 * Returns
 *   std::vector<T>
 *
 * Throws
 *   Throws std::exception if ColumnPtr instance is invalid or if data cannot be retrieved
 *   for some unknown reason. Also throws if the incorrect C++ data type is used for the
 *   column's type (see below).
 *
 * Notes
 *   Supported C++ data types: double, float, int, long, unsigned long, short, unsigned short,
 *   std::wstring, std::string, byte, char, std::complex<double>.
 *
 *   The C++ data type of the vector must match the column's type based on map below.
 *
 * Example
 *   std::vector<double> vec = orglab_data::get_column_data<double>(col); // col is Text & Numeric
 *
 */
template<class T> std::vector<T> get_column_data<T>(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1)
```

```cpp
/* Sets (inserts) matrix data using an orglab_data::matrix_adapter.
 *
 * Parameters
 *   MatrixObjectPtr                    ptr Instance representing a matrix in a worksheet.
 *   orglab_data::matrix_adapter<T>     ma  Instance of orglab_data::matrix_adapter. Supported C++ data types are listed below.
 *
 * Returns
 *   void
 *
 * Throws
 *   Throws std::exception if MatrixObjectPtr instance is invalid or if data cannot be inserted
 *   for some unknown reason.
 *
 * Notes
 *   Supported C++ data types: double, float, int, long, unsigned long, short, unsigned short,
 *   byte, char, std::complex<double> (NOT std::wstring or std::string).
 *
 *   Converts matrix type to the appropriate one mapped to the data type below. This behavior
 *   may be turned off by defining: ORGLAB_DATA_NO_CHANGE_DATA_TYPE
 */
void set_matrix_data(const MatrixObjectPtr& ptr, const orglab_data::matrix_adapter<T>& ma)
```

```cpp
/* Gets (retrieves) an orglab_data::matrix_adapter from a matrix.
 *
 * Parameters
 *   MatrixObjectPtr    ptr Instance representing a column in a worksheet.
 *
 * Returns
 *   orglab_data::matrix_adapter<T>
 *
 * Throws
 *   Throws std::exception if MatrixObjectPtr instance is invalid or if data cannot be retrieved
 *   for some unknown reason. Also throws if the incorrect C++ data type is used for the
 *   matrix's type (see below).
 *
 * Notes
 *   Supported C++ data types: double, float, int, long, unsigned long, short, unsigned short,
 *   byte, char, std::complex<double> (NOT std::wstring or std::string).
 *
 *   The C++ data type of the orglab_data::matrix_adapter must match the matrix's type based on map below.
 *
 * Example
 *   orglab_data::matrix_adapter ma = orglab_data::get_matrix_data<double>(mat); // mat is Double
 *
 */
orglab_data::matrix_adapter<T> = get_matrix_data(const MatrixObjectPtr& ptr)
```

```cpp
/* Converts a std::string or std::wstring into the appropriate type for a string-based COM object property.
 *
 * Parameters
 *   std::string or std::wstring	str 	Value to convert.
 *
 * Returns
 *   _bstr_t.
 *
 * Example
 *   col_ptr_1->LongName = orglab_data::to_str_prop("Time");
 *
 */
_bstr_t to_str_prop(const std::wstring& str)
_bstr_t to_str_prop(const std::string& str)
```

```cpp
/* Gets (retrieves) a string-based COM object property into a std::string or std::wstring.
 *
 * Parameters
 *   _bstr_t	prop 	COM object string property.
 *
 * Returns
 *   std::string or std::wstring.
 *
 * Example
 *   std::wstring ln = orglab_data::from_str_prop<std::wstring>(col_ptr_1->LongName);
 *
 */
std::wstring from_str_prop<std::wstring>(const _bstr_t& prop)
std::string from_str_prop<std::string>(const _bstr_t& prop)
```

#### Map of C++ data types to Origin types

```cpp
double					<==>	COLDATAFORMAT::DF_TEXT_NUMERIC
double					<==>	COLDATAFORMAT::DF_DOUBLE
double					<==>	COLDATAFORMAT::DF_DATE
double					<==>	COLDATAFORMAT::DF_TIME
float					<==>	COLDATAFORMAT::DF_FLOAT
int					<==>	COLDATAFORMAT::DF_LONG
long					<==>	COLDATAFORMAT::DF_LONG
unsigned long			        <==>	COLDATAFORMAT::DF_ULONG
short					<==>	COLDATAFORMAT::DF_SHORT
unsigned short			        <==>	COLDATAFORMAT::DF_USHORT
std::wstring			        <==>	COLDATAFORMAT::DF_TEXT_NUMERIC
std::string			        <==>	COLDATAFORMAT::DF_TEXT_NUMERIC
byte					<==>	COLDATAFORMAT::DF_BYTE
char					<==>	COLDATAFORMAT::DF_CHAR
std::complex<double>	                <==>	COLDATAFORMAT::DF_COMPLEX
```

#### orglab::matrix_adapter Class

Simple 2D row-major C++ "matrix" class meant to make it easier to interact with MatrixObjectPtr.
It saves a lot of coding pain.

Both the setting and getting functions for matrix data utilize an instance of the class. E.g.:

```cpp
orglab_data::matrix_adapter<double> ma_1(5, 7); // 5 rows, 7 cols
l = 1;
for (unsigned short i = 0; i < 5; ++i) {
	for (unsigned short j = 0; j < 7; ++j, ++l) {
		ma_1(i, j) = 1.0 * l;// rows, cols.
	}
}
orglab_data::set_matrix_data(mat_ptr_1, ma_1);

orglab_data::matrix_adapter<double> ma_2 = orglab_data::get_matrix_data<double>(mat_ptr_1);
```

```cpp
template<class T>
class matrix_adapter {

    /* Constructor */
    explicit matrix_adapter(const T& fill = T())

    /* Constructor */
    explicit matrix_adapter(const unsigned short& rows, const unsigned short& cols, const T& fill = T())

    /* Constructor assigns raw array copying data */
    explicit matrix_adapter(T* data, const unsigned short& rows, const unsigned short& cols)

    /* Assignment operator. Returns reference for given row and column */
    inline T& operator() (const unsigned short& row, const unsigned short& col)

    /* Read operator. Returns value for given row and column */
    inline T operator() (const unsigned short& row, const unsigned short& col) const

    /* Assigns raw array to matrix adapter */
    matrix_adapter& assign(T* data, const unsigned short& rows, const unsigned short& cols)

    /* Returns raw array of internal storage */
    const T* data() const

    /* Returns number of rows */
    inline unsigned short rows() const

    /* Returns number of columns */
    inline unsigned short cols() const

    /* Return size in elements of internal storage */
    inline std::size_t size() const

    /* Changes dimensions of matrix adapter clearing data */
    matrix_adapter& resize(const unsigned short& rows, const unsigned short& cols)

    /* Resets matrix adapter */
    matrix_adapter& clear()

    /* Transposes and returns copy of matrix adapter */
    template <class T>
    matrix_adapter transpose() const

    /* Transposes matrix adapter in place */
    matrix_adapter& transpose_self()
};
```
