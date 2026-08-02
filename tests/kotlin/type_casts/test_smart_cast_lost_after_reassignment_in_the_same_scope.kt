// vybe-test: kotlin/type_casts/test_smart_cast_lost_after_reassignment_in_the_same_scope
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            var value: Any = "start"
            if (value is String) {
                println(value.length)
                value = 9
            }
            if (value is String) {
                println("after-string")
            } else {
                println("after-not-string")
            }
        }

