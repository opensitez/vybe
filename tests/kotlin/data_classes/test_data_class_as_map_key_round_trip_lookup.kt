// vybe-test: kotlin/data_classes/test_data_class_as_map_key_round_trip_lookup
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Entry(val k: Int, val v: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf(Entry(1, 2) to "ok")
            __check((map[Entry(1, 2)]).toString(), "ok")
            __check((map[Entry(2, 1)] == null).toString(), "true")
        }
