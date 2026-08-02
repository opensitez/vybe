// vybe-test: kotlin/extension_properties/test_extension_property_pair_left
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val <A, B> Pair<A, B>.left: A get() = first
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Pair("a", 1).left).toString(), "a")
        }
