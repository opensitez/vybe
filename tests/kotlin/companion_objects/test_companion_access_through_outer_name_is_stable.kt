// vybe-test: kotlin/companion_objects/test_companion_access_through_outer_name_is_stable
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Counter {
            companion object {
                val start = 5
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter.start).toString(), "5")
            __check((Counter.Companion.start).toString(), "5")
        }
