// vybe-test: kotlin/secondary_constructors/test_secondary_with_array_input
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Bucket {
            val size: Int

            constructor(values: Array<Int>) {
                this.size = values.size
            }

            constructor(a: Int) : this(arrayOf(a)) {}
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Bucket(5).size).toString(), "1")
            __check((Bucket(arrayOf(1, 2, 3)).size).toString(), "3")
        }
