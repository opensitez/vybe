// vybe-test: kotlin/annotations/test_annotation_usage_on_local_type_alias_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class Local

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            @Local
            class Holder(val value: Int)
            val h = Holder(9)
            __check((h.value).toString(), "9")
        }
