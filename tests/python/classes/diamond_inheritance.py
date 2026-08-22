# vybe-test: python/classes/diamond_inheritance
# origin: languages/python/tests/python/test_classes.rs

class Base:
    pass
class Left(Base):
    pass
class Right(Base):
    pass
class Child(Left, Right):
    pass
