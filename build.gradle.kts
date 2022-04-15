group = "io.github.evanrupert"
version = "1.0.2"

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
    api("org.apache.poi:poi-ooxml:5.2.2")
    implementation("org.apache.logging.log4j:log4j-api:2.17.2")
    implementation("org.apache.logging.log4j:log4j-core:2.17.2")

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
        create<MavenPublication>("ossrh") {
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

    repositories {
        maven {
            val releaseUrl = "https://s01.oss.sonatype.org/service/local/staging/deploy/maven2/"
            val snapshotUrl = "https://s01.oss.sonatype.org/content/repositories/snapshots"

            url = uri(if (version.toString().endsWith("SNAPSHOT")) snapshotUrl else releaseUrl)

            credentials {
                val ossrhUsername: String by project
                val ossrhPassword: String by project

                username = ossrhUsername
                password = ossrhPassword
            }
        }
    }
}

signing {
    sign(publishing.publications["ossrh"])
}
