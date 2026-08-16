! vybe-test: fortran/interfaces/if_proc_result_12
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: v
v = f()
if (v /= 1) then
    print *, "FAIL: want [1] got [", v, "]"
    stop 1
end if
contains
function f() result(r)
integer :: r
r=1
end function f
end program t
