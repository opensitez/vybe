// vybe-test: kotlin/type_aliases/test_typealias_for_generic_factory_function
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

data class Box<T>(val value: T)

        typealias BoxFactory<T> = (T) -> Box<T>

        fun make(value: Int, factory: BoxFactory<Int>): Box<Int> {
            return factory(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val boxFactory: BoxFactory<Int> = { Box(it * 2) }
            __check((make(4, boxFactory).value).toString(), "8")
        }
