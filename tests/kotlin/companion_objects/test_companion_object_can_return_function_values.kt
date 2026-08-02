// vybe-test: kotlin/companion_objects/test_companion_object_can_return_function_values
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Math {
            companion object {
                fun build(prefix: String): (Int) -> Int {
                    return { value -> value + prefix.length }
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val add = Math.build("hello")
            __check((add(5)).toString(), "10")
        }
