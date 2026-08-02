// vybe-test: kotlin/companion_objects/test_companion_calls_from_nested_types_share_parent_state
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Board {
            companion object {
                private var tokens = 0
                fun hit(): Int {
                    tokens += 1
                    return tokens
                }
            }

            class Checker {
                fun hit(): Int = Board.hit()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Board.hit()).toString(), "1")
            __check((Board.Checker().hit()).toString(), "2")
            __check((Board.hit()).toString(), "3")
        }
