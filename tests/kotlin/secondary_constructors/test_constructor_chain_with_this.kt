// vybe-test: kotlin/secondary_constructors/test_constructor_chain_with_this
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Margin {
            val top: Int
            val right: Int
            val bottom: Int
            val left: Int

            constructor(all: Int) : this(all, all, all, all) {
                __check(("all").toString(), "all")
            }

            constructor(top: Int, right: Int, bottom: Int, left: Int) {
                this.top = top
                this.right = right
                this.bottom = bottom
                this.left = left
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Margin(7)
            __check((m.top).toString(), "7")
            __check((m.right).toString(), "7")
            __check((m.bottom).toString(), "7")
            __check((m.left).toString(), "7")
        }
