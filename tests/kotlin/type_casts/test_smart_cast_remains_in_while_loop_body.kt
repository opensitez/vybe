// vybe-test: kotlin/type_casts/test_smart_cast_remains_in_while_loop_body
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            var value: Any? = "abc"
            while (value is String) {
                println(value.length)
                value = 0
            }
        }

