// vybe-test: kotlin/collections_set/test_set_join_and_joined_string
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf(1, 2, 3)
            __check((values.joinToString(",")).toString(), "1,2,3")
            __check((values.joinToString("|") { it.toString() }).toString(), "1|2|3")
            __check((values.joinToString("") { (it * 2).toString() }).toString(), "246")
        }
