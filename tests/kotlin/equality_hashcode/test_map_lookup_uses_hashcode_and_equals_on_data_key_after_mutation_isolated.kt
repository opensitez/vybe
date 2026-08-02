// vybe-test: kotlin/equality_hashcode/test_map_lookup_uses_hashcode_and_equals_on_data_key_after_mutation_isolated
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class MutableHolder(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = MutableHolder(1)
            val map = hashMapOf(original to "start")
            original.value = 2
            __check((map.containsKey(MutableHolder(1))).toString(), "false")
            __check((map.containsKey(MutableHolder(2))).toString(), "false")
            __check((map[MutableHolder(2)]).toString(), "null")
        }
