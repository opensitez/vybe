! vybe-test: fortran/fortran2018_extended/reduce_product_with_identity_one
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: a(4) = [2, 3, 4, 5]
if ((reduce(a, operator(*), identity=1)) /= 120) then
    print *, "FAIL: want [120] got [", reduce(a, operator(*), identity=1), "]"
    stop 1
end if
end program t
