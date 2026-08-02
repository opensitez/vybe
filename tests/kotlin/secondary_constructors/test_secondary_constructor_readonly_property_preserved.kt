// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_readonly_property_preserved
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Metric {
            val total: Int
            var tag: String

            constructor(base: Int) {
                this.total = base
                this.tag = "base"
            }

            constructor(base: Int, tag: String) : this(base) {
                this.tag = tag
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = Metric(3)
            val two = Metric(5, "custom")
            __check((one.total).toString(), "3")
            __check((one.tag).toString(), "base")
            __check((two.total).toString(), "5")
            __check((two.tag).toString(), "custom")
        }
