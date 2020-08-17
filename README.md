# orglab_data

- Single file (header only library): [orglab_data.hpp](orglab_data.hpp)
- C++17 required.
- 64-bit (x64).
- Can use either Origin Automation Server or OrgLab.
- Please, please look at the [example Visual Studio 2019 solution](orglab_data_example) for good examples and additional important information.

### API

#### Setup
Prior to using the API, you need to assign the `#ORGLAB_DATA_ORGLAB_NS` macro to the name of the namespace you used when importing the type library for either Origin Automation Server or OrgLab. You HAVE to do this. For example:

```cpp
#import "Origin8.tlb" no_dual_interfaces rename_namespace("origin")
#define ORGLAB_DATA_ORGLAB_NS origin
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
template<class T> std::vector<T> get_column_data(const ColumnPtr& ptr, const long& offset = 0, const long& rows = -1)
```

#### Map of C++ data types to column types

```cpp
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
```
