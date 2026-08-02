// vybe-test: kotlin/equality_hashcode/test_hashcode_matches_structural_equality
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(val a: Int, val b: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Item(9, "ok")
            val right = Item(9, "ok")
            __check((left.hashCode() == right.hashCode()).toString(), "true")
        }
