// vybe-test: kotlin/secondary_constructors/test_secondary_nullable_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Value {
            val value: Int?

            constructor(v: Int) {
                this.value = v
            }

            constructor(flag: Boolean) : this(if (flag) 1 else 0)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Value(true).value).toString(), "1")
            __check((Value(false).value).toString(), "0")
        }
