! vybe-test: fortran/pack_unpack_extended/unpack_int_alternating_restore
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [10, 30, 50]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: fill(5) = [0, 0, 0, 0, 0]
integer :: b(5)
b = unpack(a, mask, fill)
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
if ((b(5)) /= 50) then
    print *, "FAIL: want [50] got [", b(5), "]"
    stop 1
end if
end program t
