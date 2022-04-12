group = "io.github.evanrupert"
version = "1.0.0"

plugins {
    // Apply the Kotlin JVM plugin to add support for Kotlin.
    kotlin("jvm") version "1.6.20"
    id("org.jetbrains.dokka") version "1.6.10"

    // Apply the java-library plugin for API and implementation separation.
    `java-library`

    // Apply maven-publish plugin to generate artifacts to upload to maven central
    `maven-publish`

    signing
}

repositories {
    mavenCentral()
}

dependencies {
    // Apache POI Excel library
    implementation("org.apache.poi:poi-ooxml:3.9")

    // Use the Kotlin test library.
    testImplementation("org.jetbrains.kotlin:kotlin-test")

    // Use the Kotlin JUnit integration.
    testImplementation("org.jetbrains.kotlin:kotlin-test-junit")

    // Use the strikt assertion library
    testImplementation("io.strikt:strikt-core:0.34.0")

    // Use mockito for mocking apache poi in unit tests
    testImplementation("org.mockito.kotlin:mockito-kotlin:4.0.0")
}

val dokkaJavadocJar by tasks.register<Jar>("dokkaJavadocJar") {
    dependsOn(tasks.dokkaJavadoc)
    from(tasks.dokkaJavadoc.flatMap { it.outputDirectory })
    archiveClassifier.set("javadoc")
}

val dokkaHtmlJar by tasks.register<Jar>("dokkaHtmlJar") {
    dependsOn(tasks.dokkaHtml)
    from(tasks.dokkaHtml.flatMap { it.outputDirectory })
    archiveClassifier.set("html-doc")
}

java {
    withSourcesJar()
}

publishing {
    publications {
        create<MavenPublication>("library") {
            from(components["java"])

            artifact(dokkaJavadocJar)
            artifact(dokkaHtmlJar)

            pom {
                name.set("ExcelKt")
                description.set("Kotlin Wrapper over the Apache POI Excel Library that enables creating xlsx files with kotlin builders")
                url.set("https://github.com/evanrupert/ExcelKt")

                licenses {
                    license {
                        name.set("MIT")
                        url.set("https://opensource.org/licenses/MIT")
                    }
                }
                developers {
                    developer {
                        id.set("evanrupert")
                        name.set("Evan Rupert")
                        email.set("rupertevanr@gmail.com")
                        url.set("https://github.com/evanrupert")
                    }
                }
                scm {
                    url.set("https://github.com/evanrupert/ExcelKt")
                    connection.set("scm:git:git://github.com/evanrupert/ExcelKt.git")
                    developerConnection.set("scm:git:git@github.com:evanrupert/ExcelKt.git")
                }
            }
        }
    }
}

signing {
    sign(publishing.publications["library"])
}
