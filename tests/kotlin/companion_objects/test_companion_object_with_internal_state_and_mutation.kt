// vybe-test: kotlin/companion_objects/test_companion_object_with_internal_state_and_mutation
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Store {
            companion object {
                private var next: Int = 0
                fun take(): Int {
                    next += 1
                    return next
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
            __check((Store.take()).toString(), "1")
            __check((Store.take()).toString(), "2")
        }
