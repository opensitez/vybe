// vybe-test: kotlin/kotlin_type_checks/test_smart_cast_after_type_check
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_checks.rs

fun main() {
            val value: Any = 12L
            if (value is Long) {
                println(value + 3)
            } else {
                println(0)
            }
        }

