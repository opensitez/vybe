! vybe-test: fortran/modulo_dim_sign_extended/dim_in_sum_accumulator
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
integer :: a(5)=[3,8,1,9,4]
integer :: b(5)=[7,2,6,1,5]
if ((sum(dim(a,b))) /= 14) then
    print *, "FAIL: want [14] got [", sum(dim(a,b)), "]"
    stop 1
end if
end program t
