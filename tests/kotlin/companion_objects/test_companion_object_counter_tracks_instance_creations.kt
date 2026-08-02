// vybe-test: kotlin/companion_objects/test_companion_object_counter_tracks_instance_creations
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Token {
            companion object {
                var total = 0
            }

            init {
                Token.total += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Token()
            Token()
            Token()
            __check((Token.total).toString(), "3")
        }
