// vybe-test: kotlin/smart_casts/test_as_cast_failure_throws_class_cast_exception
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun main() {
            try {
                val value: Any = 7
                val text = value as String
                println(text)
            } catch (error: ClassCastException) {
                println("cast-failed")
            }
        }

