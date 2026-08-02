// vybe-test: kotlin/interfaces/test_interface_object_expression_mutating_capture
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Mutator {
            fun next(): Int
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var n = 1
            val m: Mutator = object : Mutator {
                override fun next(): Int {
                    val out = n
                    n += 1
                    return out
                }
            }
            __check((m.next()).toString(), "1")
            __check((m.next()).toString(), "2")
        }
