! vybe-test: fortran/alternate_returns/alternate_returns_08
! origin: languages/fortran/tests/fortran/test_alternate_returns.rs
subroutine s(x,*,*)
integer::x
return 2
end
program driver
integer :: n
n = 0
call s(5, *10, *20)
n = 99
go to 90
10 n = 1
go to 90
20 n = 2
go to 90
90 continue
if (n /= 2) then
    print *, "FAIL: want [2] got [", n, "]"
    stop 1
end if
end program driver
