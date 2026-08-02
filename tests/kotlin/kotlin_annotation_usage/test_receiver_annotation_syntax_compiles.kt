// vybe-test: kotlin/kotlin_annotation_usage/test_receiver_annotation_syntax_compiles
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.RECEIVER)
        annotation class Ext

        class Token

        @Ext
        fun Token.wrap(prefix: String) = prefix + "#"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Token().wrap("x")).toString(), "x#")
        }
