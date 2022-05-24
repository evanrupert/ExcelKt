ExcelKt
-
An idiomatic Kotlin wrapper over the Apache POI Excel library for easily generating Excel xlsx files.

### Key Features
- Write idiomatic kotlin that looks clean and logical
- Simple style system that allows you to use the Apache POI CellStyle system to stylize workbooks, sheets, rows, or specific cells
- Very lightweight

### Installation

#### Gradle

In your gradle build file add the following:

Kotlin DSL
```kotlin
repositories {
    mavenCentral()
}

dependencies {
    implementation("io.github.evanrupert:excelkt:1.0.2")
}
```

Groovy DSL
```groovy
repositories {
    mavenCentral()
}

dependencies {
    implementation 'io.github.evanrupert:excelkt:1.0.2'
}
```

#### Maven

In your `pom.xml` file make sure you have the following in your `repositories`:
```xml
<repository>
    <id>mavenCentral</id>
    <url>https://repo1.maven.org/maven2/</url>
</repository>
```

Then add the following to your `dependencies`:
```xml
<dependency>
    <groupId>io.github.evanrupert</groupId>
    <artifactId>excelkt</artifactId>
    <version>1.0.2</version>
</dependency>
```

#### Legacy Installation

For older versions of ExcelKt (`v0.1.2` and before), which run on kotlin version `1.3.x` and apache poi version `3.9`, add the following to your gradle build file:

Kotlin DSL
```kotlin
repositories {
    jcenter()
    maven(url = "https://jitpack.io")
}

dependencies {
    implementation("com.github.EvanRupert:ExcelKt:v0.1.2")
}
```

Groovy DSL
```groovy
repositories {
    jcenter()
    maven { url 'https://jitpack.io' }
}

dependencies {
    implementation 'com.github.evanrupert:excelkt:v0.1.2'
}
```

And use `import excelkt.*` instead of `import io.github.evanrupert.excelkt.*`

### Quick Example
```kotlin
import io.github.evanrupert.excelkt.*
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors

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

        fillPattern = FillPatternType.SOLID_FOREGROUND
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
