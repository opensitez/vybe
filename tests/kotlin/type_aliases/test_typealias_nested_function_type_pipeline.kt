// vybe-test: kotlin/type_aliases/test_typealias_nested_function_type_pipeline
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Transformer<T> = (T) -> T
        typealias IntTransformer = Transformer<Int>

        fun applyTwice(value: Int, first: IntTransformer, second: IntTransformer): Int {
            return second(first(value))
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stepA: IntTransformer = { it + 5 }
            val stepB: IntTransformer = { it * 2 }
            __check((applyTwice(3, stepA, stepB)).toString(), "16")
        }
