// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_vararg_forwarding
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class VarArgProbe {
            val text: String

            constructor(prefix: String, vararg values: Int) {
                this.text = prefix + values.joinToString(":")
            }

            constructor(value: Int) : this("n", value, value + 1) {}
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((VarArgProbe("v", 1, 2, 3).text).toString(), "v:1:2:3")
            __check((VarArgProbe(4).text).toString(), "n:4:5")
        }
