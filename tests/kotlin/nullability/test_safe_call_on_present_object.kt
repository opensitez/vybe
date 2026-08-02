// vybe-test: kotlin/nullability/test_safe_call_on_present_object
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Item(val price: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Item? = Item(49)
            __check((item?.price ?: 0).toString(), "49")
        }
