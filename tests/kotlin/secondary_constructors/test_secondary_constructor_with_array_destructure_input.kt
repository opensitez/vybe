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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Matrix(2, 3)
            val b = Matrix(arrayOf(arrayOf(1, 2), arrayOf(3, 4), arrayOf(5, 6)))
            __p((a.rows).toString())
            __p((a.cols).toString())
            __p((b.rows).toString())
            __p((b.cols).toString())
        
__check("2\n3\n3\n2")
}
