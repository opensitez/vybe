// vybe-test: kotlin/secondary_constructors/test_secondary_nested_class_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Holder {
            val x: Int
            class Child

            constructor() {
                this.x = 1
            }

            constructor(v: Int) : this() {
                this.x = v
            }

            constructor(v: Int, child: Child) : this(v) {
                child
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder(7, Holder.Child())
            __check((h.x).toString(), "7")
        }
