// vybe-test: kotlin/generics/test_generic_typealias_builder_infers_type
// origin: languages/kotlin/tests/kotlin/test_generics.rs

typealias Factory<T> = () -> T

        fun <T> materialize(factory: Factory<T>): T {
            return factory()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val number = materialize { 9 }
            val text = materialize { "go" }
            __check((number).toString(), "9")
            __check((text).toString(), "go")
        }
