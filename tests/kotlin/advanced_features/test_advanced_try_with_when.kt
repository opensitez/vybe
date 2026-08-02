// vybe-test: kotlin/advanced_features/test_advanced_try_with_when
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun main() { try { println("ok") } catch (e: Exception) { println("bad") } finally { println("end") } }

