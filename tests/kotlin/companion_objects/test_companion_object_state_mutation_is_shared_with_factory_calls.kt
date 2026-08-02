// vybe-test: kotlin/companion_objects/test_companion_object_state_mutation_is_shared_with_factory_calls
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Counter {
            companion object {
                private var next = 0

                fun next(delta: Int = 1): Int {
                    next += delta
                    return next
                }

                fun current(): Int = next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter.next()).toString(), "1")
            __check((Counter.next(3)).toString(), "4")
            __check((Counter.next()).toString(), "5")
            __check((Counter.current()).toString(), "5")
        }
