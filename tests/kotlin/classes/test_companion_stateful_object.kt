// vybe-test: kotlin/classes/test_companion_stateful_object
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Holder {
            companion object {
                var created = 0
                fun create(): Holder {
                    created += 1
                    return Holder()
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
            Holder.create()
            Holder.create()
            __check((Holder.created).toString(), "2")
        }
