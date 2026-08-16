! vybe-test: fortran/interface_blocks/interface_blocks_07
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
subroutine s(a)
real::a(:)
a = a * 2.0
end subroutine s
program t

interface
subroutine s(a)
real::a(:)
end subroutine s
end interface
real :: buf(3)
buf = [1.0, 2.0, 3.0]
call s(buf)
if (nint(sum(buf)) /= 12) then
    print *, "FAIL: want [12] got [", nint(sum(buf)), "]"
    stop 1
end if
end program t
