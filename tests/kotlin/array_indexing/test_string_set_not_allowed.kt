// vybe-test: kotlin/array_indexing/test_string_set_not_allowed
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun main() {
        val s = "kotlin"
        try {
            // not actually executable by design
            val value = s[0]
            println(value)
        } catch (e: Exception) {
            println("ok")
        }
    }

