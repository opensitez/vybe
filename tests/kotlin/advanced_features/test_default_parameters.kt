// vybe-test: kotlin/advanced_features/test_default_parameters
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun greet(name: String = "World") {
            println("Hello " + name)
        }

        fun main() {
            greet()
            greet("Kotlin")
        }

