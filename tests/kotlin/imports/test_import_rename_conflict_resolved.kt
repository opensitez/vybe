// vybe-test: kotlin/imports/test_import_rename_conflict_resolved
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.setOf as setOfStrings
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = setOfStrings("a", "b")
            __check((s.size).toString(), "2")
        }
