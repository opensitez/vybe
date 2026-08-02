// vybe-test: kotlin/generics/test_generic_extension_constraint_and_mapping
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Wrapper<T>(val value: T)

        fun <T> Wrapper<T>.map(transform: (T) -> T): T {
            return transform(this.value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Wrapper("a").map { it + "b" }).toString(), "ab")
            __check((Wrapper(9).map { it + 1 }).toString(), "10")
        }
