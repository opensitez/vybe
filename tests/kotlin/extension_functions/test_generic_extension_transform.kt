// vybe-test: kotlin/extension_functions/test_generic_extension_transform
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun <T> List<T>.wrapCount(): String = "count=" + this.size

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf(1, 2, 3).wrapCount()).toString(), "count=3")
            __check((listOf("a").wrapCount()).toString(), "count=1")
        }
