// vybe-test: kotlin/kotlin_companion_objects_api/test_companion_access_as_property_like_accessor
// origin: languages/kotlin/tests/kotlin/test_kotlin_companion_objects_api.rs

class Counter {
            companion object {
                var current: Int = 0

                fun bump(): Int {
                    current += 1
                    return current
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
            __check((Counter.bump()).toString(), "1")
            __check((Counter.current).toString(), "1")
        }
