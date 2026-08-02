// vybe-test: kotlin/kotlin_iterable_to_collections/test_associate_by_to
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("aa", "bb", "c")
            val map = linkedMapOf<Int, String>()
            values.associateByTo(map) { it.length }
            __check((map.keys.joinToString(",")).toString(), "2,1")
            __check((map[2]).toString(), "bb")
        }
