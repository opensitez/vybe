// vybe-test: kotlin/lambdas/test_lambda_parameter_named_inference_and_type_annotation
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun <T> apply(values: List<T>, transform: (T) -> Int): Int {
    return values.map(transform).fold(0) { left, right -> left + right }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    __check((apply(listOf("aa", "b", "ccc"), { it.length })).toString(), "6")
}
