// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_super_argument_is_evaluated_for_derived_init
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

var calls = ""

        open class Parent(val seed: Int) {
            val value = seed + 1
        }

        fun computeSeed(base: Int): Int {
            calls += base.toString()
            return base * 10
        }

        class Child : Parent {
            val offset: Int

            constructor(base: Int) : super(computeSeed(base)) {
                this.offset = base + 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Child(7)
            __check((c.value).toString(), "71")
            __check((c.offset).toString(), "8")
            __check((calls).toString(), "7")
        }
