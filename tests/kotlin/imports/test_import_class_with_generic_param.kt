// vybe-test: kotlin/imports/test_import_class_with_generic_param
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.ArrayList
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list: ArrayList<Int> = ArrayList()
            list.add(1)
            list.add(2)
            __check((list[0] + list[1]).toString(), "3")
        }
