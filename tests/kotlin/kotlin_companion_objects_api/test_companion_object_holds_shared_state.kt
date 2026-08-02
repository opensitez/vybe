// vybe-test: kotlin/kotlin_companion_objects_api/test_companion_object_holds_shared_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_companion_objects_api.rs

class Sequence {
            private val id: Int

            private constructor(value: Int) {
                id = value
            }

            companion object {
                var next: Int = 0
                fun nextSequence(): Sequence {
                    next += 1
                    return Sequence(next)
                }
            }

            fun value(): Int = id
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Sequence.nextSequence()
            val b = Sequence.nextSequence()
            __check((a.value()).toString(), "1")
            __check((b.value()).toString(), "2")
            __check((Sequence.next).toString(), "2")
        }
