// vybe-test: kotlin/kotlin_iterable_to_collections/test_associate_with_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("x", "y", "z")
            val out = values.associateWith { it.length + 1 }
            __check((out.size).toString(), "3")
            __check((out["y"].toString()).toString(), "2")
        }
