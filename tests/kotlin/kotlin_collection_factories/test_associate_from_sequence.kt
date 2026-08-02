// vybe-test: kotlin/kotlin_collection_factories/test_associate_from_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = sequenceOf("a", "bb", "ccc").associateWith { it.length }
            __check((values["a"]).toString(), "1")
            __check((values["ccc"]).toString(), "3")
        }
