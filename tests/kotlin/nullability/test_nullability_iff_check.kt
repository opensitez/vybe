// vybe-test: kotlin/nullability/test_nullability_iff_check
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun main() { val value: Int? = 10
if (value != null) { println(value) } else { println(0) } }

