// vybe-test: kotlin/object_expressions/test_object_expression_as_argument
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Sink { fun consume(v: Int) }
fun call(sink: Sink, value: Int) = sink.consume(value)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { call(object : Sink { override fun consume(v: Int) { __check((v).toString(), "7") } }, 7) }
