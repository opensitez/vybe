// vybe-test: kotlin/secondary_constructors/test_secondary_double_constructor
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class B {
            val a: Int
            val b: Int

            constructor(a: Int) {
                this.a = a
                this.b = a
            }

            constructor(a: Int, b: Int) : this(a) {
                this.b = b
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = B(2, 9)
            __check((x.a).toString(), "2")
            __check((x.b).toString(), "9")
        }
