// vybe-test: kotlin/kotlin_annotation_usage/test_nested_annotation_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

enum class Kind { ALPHA, BETA }
        annotation class Meta(val kind: Kind)
        annotation class Bundle(val metas: Array<Meta>)

        @Bundle([Meta(Kind.ALPHA), Meta(Kind.BETA)])
        class Combined

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Combined::class.simpleName).toString(), "Combined")
        }
