// vybe-test: kotlin/type_casts/test_mutable_list_cast_to_readonly_list_and_back
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mutable = mutableListOf(1, 2, 3)
            val asReadOnly = mutable as List<Int>
            __check((asReadOnly.size).toString(), "3")

            val backToMutable = mutable as? MutableList<Int>
            __check((backToMutable != null).toString(), "true")
            __check((backToMutable?.size ?: -1).toString(), "3")
        }
