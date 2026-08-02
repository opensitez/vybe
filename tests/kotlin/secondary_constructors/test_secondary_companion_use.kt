// vybe-test: kotlin/secondary_constructors/test_secondary_companion_use
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Factory {
            val size: Int
            companion object {
                fun create(v: Int): Int {
                    return v * 2
                }
            }

            constructor(v: Int) {
                this.size = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Factory.create(4)).toString(), "8")
            __check((Factory.create(Factory(6).size)).toString(), "12")
        }
