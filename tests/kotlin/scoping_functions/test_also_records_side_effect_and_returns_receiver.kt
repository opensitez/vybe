// vybe-test: kotlin/scoping_functions/test_also_records_side_effect_and_returns_receiver
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val events = mutableListOf<String>()
            val values = mutableListOf(1, 2).also {
                events.add("size-" + it.size.toString())
                it.add(3)
            }
            __check((values.joinToString(",")).toString(), "1,2,3")
            __check((events.joinToString("|")).toString(), "size-2")
        }
