// vybe-test: kotlin/generics/test_generic_extension_function
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<T>(val value: T)

        fun <T> Holder<T>.labeled(prefix: String): String {
            return prefix + ":" + this.value.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder(5).labeled("x")).toString(), "x:5")
            __check((Holder("one").labeled("y")).toString(), "y:one")
        }
