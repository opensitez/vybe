# vybe-test: python/decimal_fractions_cmath/fractions_from_decimal
# origin: languages/python/tests/python/test_decimal_fractions_cmath.rs
# vybe-test-mode: compile

from fractions import Fraction
from decimal import Decimal
Fraction(Decimal('0.5'))
