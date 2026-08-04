! vybe-test: fortran/modulo_dim_sign_extended/sign_array_elementwise
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
integer :: a(3)=[5,-5,0]
integer :: s(3)=[-1,1,-1]
if ((sign(a(1), s(1))) /= -5) then
    print *, "FAIL: want [-5] got [", sign(a(1), s(1)), "]"
    stop 1
end if
if ((sign(a(2), s(2))) /= 5) then
    print *, "FAIL: want [5] got [", sign(a(2), s(2)), "]"
    stop 1
end if
if ((sign(a(3), s(3))) /= 0) then
    print *, "FAIL: want [0] got [", sign(a(3), s(3)), "]"
    stop 1
end if
end program t
