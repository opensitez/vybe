// vybe-test: kotlin/classes/test_companion_object_with_var
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Counter {
            companion object {
                var created = 0
                fun track(): Int {
                    created += 1
                    return created
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
            __check((Counter.track()).toString(), "1")
            __check((Counter.track()).toString(), "2")
        }
