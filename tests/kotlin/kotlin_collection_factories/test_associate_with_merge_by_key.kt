// vybe-test: kotlin/kotlin_collection_factories/test_associate_with_merge_by_key
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("ax", "ay", "bx").associate { it[0].toString() to it }
            __check((values["a"]).toString(), "ay")
            __check((values["b"]).toString(), "bx")
        }
