// vybe-test: kotlin/object_expressions/test_anonymous_object_expression
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val runner = object {
                fun run() {
                    __check(("anonymous object running").toString(), "anonymous object running")
                }
            }
            runner.run()
        }
