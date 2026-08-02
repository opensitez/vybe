// vybe-test: kotlin/secondary_constructors/test_constructor_default_behavior
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Timer {
            val seconds: Int

            constructor() {
                this.seconds = 0
            }

            constructor(start: Int) {
                this.seconds = start
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Timer().seconds).toString(), "0")
            __check((Timer(30).seconds).toString(), "30")
        }
