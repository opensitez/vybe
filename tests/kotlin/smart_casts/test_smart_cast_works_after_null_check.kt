// vybe-test: kotlin/smart_casts/test_smart_cast_works_after_null_check
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun main() {
            val value: Any? = "x"
            if (value != null) {
                println(value is String)
                println(value.uppercase())
            } else {
                println(false)
            }
        }

