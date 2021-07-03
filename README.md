ExcelKt
-
An idiomatic Kotlin wrapper over the Apache POI Excel library for easily generating Excel xlsx files.

### Key Features
- Write idiomatic kotlin that looks clean and logical
- Simple style system that allows you to use the Apache POI CellStyle system to stylize workbooks, sheets, rows, or specific cells
- Very lightweight

### Installation
In your `build.gradle.kts` or `build.gradle` add the following

If `build.gradle.kts`
```kotlin
repositories {
    jcenter()
    maven(url = "https://jitpack.io")
}

dependencies {
    implementation("com.github.EvanRupert:ExcelKt:v0.1.2")
}
```
If `build.gradle`
```groovy
repositories {
    jcenter()
    maven { url 'https://jitpack.io' }
}

dependencies {
    implementation 'com.github.EvanRupert:ExcelKt:v0.1.2'
}
```

### Quick Example
```kotlin
data class Customer(
    val id: String,
    val name: String,
    val address: String,
    val age: Int
)

fun findCustomers(): List<Customer> = listOf(
    Customer("1", "Robert", "New York", 32),
    Customer("2", "Bobby", "Florida", 12)
)

fun main() {
    workbook {
        sheet {
            row {
                cell("Hello, World!")
            }
        }

        sheet("Customers") {
            customersHeader()

            for (customer in findCustomers())
                row {
                    cell(customer.id)
                    cell(customer.name)
                    cell(customer.address)
                    cell(customer.age)
                }
        }

    }.write("test.xlsx")
}

fun Sheet.customersHeader() {
    val headings = listOf("Id", "Name", "Address", "Age")

    val headingStyle = createCellStyle {
        setFont(createFont {
            fontName = "IMPACT"
            color = IndexedColors.PINK.index
        })

        fillPattern = XSSFCellStyle.SOLID_FOREGROUND
        fillForegroundColor = IndexedColors.AQUA.index
    }

    row(headingStyle) {
        headings.forEach { cell(it) }
    }
}
```
### Supported cell data types
Cells support the following content types:
- Formula
- Boolean
- Number
- Date
- Calendar
- LocalDate
- LocalDateTime

All other types will be converted to strings.

Example of all data types in use:
```kotlin
row {
    cell(Formula("A1 + A2"))
    cell(true)
    cell(12.2)
    cell(Date())
    cell(Calendar.getInstance())
    cell(LocalDate.now())
    cell(LocalDateTime.now())
}
```

### Note on Dates
By default, dates will display as numbers in Excel.  In order to display them correctly, create a cell style with the `dataFormat` set to your preferred format.  See the following example:
```kotlin
row {
    val cellStyle = createCellStyle {
        dataFormat = xssfWorkbook.creationHelper.createDataFormat().getFormat("m/d/yy h:mm")
    }

    cell(Date(), cellStyle)
}
```
