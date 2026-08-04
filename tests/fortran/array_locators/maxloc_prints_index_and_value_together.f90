! vybe-test: fortran/array_locators/maxloc_prints_index_and_value_together
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [3, 1, 9, 1, 5]
integer :: loc(1)
loc = maxloc(a)
print *, loc(1), a(loc(1))
end program t
