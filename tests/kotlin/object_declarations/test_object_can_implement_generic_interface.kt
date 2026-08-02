// vybe-test: kotlin/object_declarations/test_object_can_implement_generic_interface
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Transformer<T, U> {
            fun map(value: T): U
        }

        object IntToText : Transformer<Int, String> {
            override fun map(value: Int): String = "v" + value
        }

        fun emit(transformer: Transformer<Int, String>, value: Int): String {
            return transformer.map(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((emit(IntToText, 3)).toString(), "v3")
        }
