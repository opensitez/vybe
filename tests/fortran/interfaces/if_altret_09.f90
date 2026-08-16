! vybe-test: fortran/interfaces/if_altret_09
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(*,*)
return 1
end
program t
integer :: n
external s
n = 0
call s(*10, *20)
n = 1
go to 30
10 n = 2
go to 30
20 n = 3
30 continue
if (n /= 2) then
    print *, "FAIL: want [2] got [", n, "]"
    stop 1
end if
end program t
