# vybe-test: python/decimal_fractions_cmath/decimal_localcontext
# origin: languages/python/tests/python/test_decimal_fractions_cmath.rs
# vybe-test-mode: compile

from decimal import localcontext, Decimal
with localcontext() as ctx:
 Decimal('1')
