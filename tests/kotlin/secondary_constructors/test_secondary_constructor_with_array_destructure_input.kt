// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_with_array_destructure_input
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Matrix {
            val rows: Int
            val cols: Int

            constructor(values: Array<Array<Int>>) {
                this.rows = values.size
                this.cols = if (values.isNotEmpty()) values[0].size else 0
            }

            constructor(rows: Int, cols: Int) : this(Array(rows) { Array(cols) { 0 } })
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Matrix(2, 3)
            val b = Matrix(arrayOf(arrayOf(1, 2), arrayOf(3, 4), arrayOf(5, 6)))
            __check((a.rows).toString(), "2")
            __check((a.cols).toString(), "3")
            __check((b.rows).toString(), "3")
            __check((b.cols).toString(), "2")
        }
