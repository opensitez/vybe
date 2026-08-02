// vybe-test: kotlin/data_class_copying/test_data_class_copy_replaces_multiple_properties
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Box(val id: Int, val tag: String, val active: Boolean)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Box(1, "x", false)
            val b = a.copy(id = 2, active = true)
            __check((a.id).toString(), "1")
            __check((b.id).toString(), "2")
            __check((b.tag).toString(), "x")
            __check((b.active).toString(), "true")
        }
