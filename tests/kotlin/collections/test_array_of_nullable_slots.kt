// vybe-test: kotlin/collections/test_array_of_nullable_slots
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val slots: Array<Int?> = arrayOf(null, null, null)
            slots[1] = 7
            __check((slots[0] == null).toString(), "true")
            __check((slots[1] + 1).toString(), "8")
            __check((slots[2] == null).toString(), "true")
        }
