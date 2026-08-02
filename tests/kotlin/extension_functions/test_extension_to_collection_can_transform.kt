// vybe-test: kotlin/extension_functions/test_extension_to_collection_can_transform
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun <T> Collection<T>.asTagged(tag: String): String {
            return tag + ":" + this.size
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf(1, 2, 3).asTagged("count")).toString(), "count:3")
            __check((setOf("a").asTagged("single")).toString(), "single:1")
        }
