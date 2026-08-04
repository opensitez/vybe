! vybe-test: fortran/legacy_data_extended/common_array_three_sum
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: arr(3)
common /nums/ arr
arr(1) = 4; arr(2) = 5; arr(3) = 6
if ((sum(arr)) /= 15) then
    print *, "FAIL: want [15] got [", sum(arr), "]"
    stop 1
end if
end program t
