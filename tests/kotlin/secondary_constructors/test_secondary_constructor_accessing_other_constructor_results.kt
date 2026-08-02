// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_accessing_other_constructor_results
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Range {
            val min: Int
            val max: Int

            constructor(value: Int) {
                this.min = value
                this.max = value
            }

            constructor(from: Int, to: Int) : this(from) {
                this.max = to
            }

            fun width(): Int = max - min
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Range(4)
            val b = Range(2, 8)
            __check((a.width()).toString(), "0")
            __check((b.width()).toString(), "6")
        }
