# vybe-test: python/numeric_semantics_extended/fractions_limit
# origin: languages/python/tests/python/test_numeric_semantics_extended.rs
# vybe-test-mode: compile

from fractions import Fraction
Fraction(1, 3) + Fraction(1, 6)
