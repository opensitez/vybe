! vybe-test: fortran/fortran2008/pack_unpack_runtime
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program t
    integer :: a(5) = [1, 2, 3, 4, 5]
    logical :: mask(5) = [.true., .false., .true., .false., .true.]
    integer :: b(3)
    integer :: c(5)
    integer :: fill(5) = [0, 0, 0, 0, 0]
    b = pack(a, mask)
    c = unpack(b, mask, fill)
    if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 3) then
    print *, "FAIL: want [3] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 5) then
    print *, "FAIL: want [5] got [", b(3), "]"
    stop 1
end if
    if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
    if ((c(5)) /= 5) then
    print *, "FAIL: want [5] got [", c(5), "]"
    stop 1
end if
end program t
