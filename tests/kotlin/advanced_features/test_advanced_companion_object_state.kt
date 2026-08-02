// vybe-test: kotlin/advanced_features/test_advanced_companion_object_state
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

class Counter {
            companion object {
                var hits: Int = 0

                fun next(): Int {
                    hits += 1
                    return hits
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
            __check((Counter.next()).toString(), "1")
            __check((Counter.next()).toString(), "2")
            __check((Counter.hits).toString(), "2")
        }
