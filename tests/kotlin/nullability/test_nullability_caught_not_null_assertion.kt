// vybe-test: kotlin/nullability/test_nullability_caught_not_null_assertion
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun main() { val value: String? = null
try { println(value!!) } catch (e: Exception) { println("fail") } }

