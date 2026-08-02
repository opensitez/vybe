# vybe-test: python/decimal_fractions_cmath/decimal_compare_total
# origin: languages/python/tests/python/test_decimal_fractions_cmath.rs
# vybe-test-mode: compile

from decimal import Decimal
Decimal('1.0').compare_total(Decimal('1'))
