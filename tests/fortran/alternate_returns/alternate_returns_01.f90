! vybe-test: fortran/alternate_returns/alternate_returns_01
! origin: languages/fortran/tests/fortran/test_alternate_returns.rs
subroutine s(*,*)
return 1
end
program driver
integer :: n
n = 0
call s(*10, *20)
n = 99
go to 90
10 n = 1
go to 90
20 n = 2
go to 90
90 continue
if (n /= 1) then
    print *, "FAIL: want [1] got [", n, "]"
    stop 1
end if
end program driver
