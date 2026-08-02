// vybe-test: kotlin/companion_objects/test_companion_object_uses_named_instance_reference
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Counter {
            companion object Factory {
                private var next = 0

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
            val first = Counter.Factory.take()
            val second = Counter.take()
            __check((first).toString(), "1")
            __check((second).toString(), "2")
        }
