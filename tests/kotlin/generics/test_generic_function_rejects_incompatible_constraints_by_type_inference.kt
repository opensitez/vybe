// vybe-test: kotlin/generics/test_generic_function_rejects_incompatible_constraints_by_type_inference
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> pairSize(left: T, right: T): Int {
            return 2
        }

        class Item

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pairSize(Item(), Item())).toString(), "2")
            val left = Item()
            val right = Item()
            __check((pairSize(left, right)).toString(), "2")
        }
