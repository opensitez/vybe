// vybe-test: kotlin/kotlin_class_init_sequences/test_companion_object_called_from_constructor
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Id {
            val value: Int
            init {
                value = next()
            }

            companion object {
                private var seq = 0
                fun next(): Int {
                    seq += 1
                    return seq
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
            __check((Id().value).toString(), "1")
            __check((Id().value).toString(), "2")
        }
