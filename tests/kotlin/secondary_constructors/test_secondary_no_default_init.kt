// vybe-test: kotlin/secondary_constructors/test_secondary_no_default_init
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class A {
            val value: Int
            constructor(v: Int) {
                this.value = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((A(3).value).toString(), "3")
        }
