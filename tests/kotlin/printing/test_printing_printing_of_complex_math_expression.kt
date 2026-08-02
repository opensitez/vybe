// vybe-test: kotlin/printing/test_printing_printing_of_complex_math_expression
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val width = 3
            val height = 4
            __check(("area=${width * height}").toString(), "area=12")
            __check(("perimeter=${2 * (width + height)}").toString(), "perimeter=14")
        }
