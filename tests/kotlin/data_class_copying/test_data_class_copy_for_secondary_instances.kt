// vybe-test: kotlin/data_class_copying/test_data_class_copy_for_secondary_instances
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Version(val major: Int, val minor: Int, val patch: Int)
        fun bumpPatch(v: Version): Version = v.copy(patch = v.patch + 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = Version(1, 2, 3)
            val n = bumpPatch(v)
            __check((n.major).toString(), "1")
            __check((n.minor).toString(), "2")
            __check((n.patch).toString(), "4")
        }
