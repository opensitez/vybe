// vybe-test: kotlin/equality_hashcode/test_hashcode_matches_equals_contract_for_class
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(val id: Int, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Item(1, "a")
            val right = Item(1, "a")
            val set = setOf(left, right)
            __check((set.size).toString(), "1")
            __check((left.hashCode() == right.hashCode()).toString(), "true")
        }
