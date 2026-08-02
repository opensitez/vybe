// vybe-test: kotlin/companion_objects/test_companion_object_isolated_state_per_host_type
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Left {
            companion object {
                var value = 1
            }
        }

        class Right {
            companion object {
                var value = 10
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Left.value += 1
            Right.value += 5
            __check((Left.value).toString(), "2")
            __check((Right.value).toString(), "15")
        }
