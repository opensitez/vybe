// vybe-test: kotlin/kotlin_multiline_strings/test_multiline_expression_embedded_interpolation
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 2
            val message = """
${'$'}n squared is ${'$'}{n * n}
${'$'}n cubed is ${'$'}{n * n * n}
"""
            val lines = message.trim().lines()
            __check((lines[0]).toString(), "2 squared is 4")
            __check((lines[1]).toString(), "2 cubed is 8")
        }
