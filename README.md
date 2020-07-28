# LockIfOutOfLogonHours - Lock Windows User Desktop When Logon Hours Denied

LockIfOutOfLogonHours is a small utility to test if the invoking user is within permitted logon hours or not.

If not, the utility locks the user desktop, locking the user out but not killing users' work.

If logon hours are to expire within 10 minutes, a prompt is displayed encouraging the user to finish her/his work in time.

## Running

Use Task Scheduler to schedule the `LockIfOutOfLogonHours.exe` to run periodically as the interactive user.

Active Directory allows configuring logon hours rounded to the whole hours only. Therefore, it makes sense to trigger the `LockIfOutOfLogonHours.exe` to run:

- Sometime before the hour is out (up to 10 minutes before the hour is out, e.g. 0:55, 1:55, ..., 23:55). This trigger notifies the user.

- At or just after the hour is out (e.g. 0:01, 1:01, ..., 23:01). This trigger locks the user out.

A sample task XML to be imported in Task Scheduler can be found in [`samples\LockIfOutOfLogonHours.xml`](samples/LockIfOutOfLogonHours.xml).

## Building

Microsoft Visual Studio 2019 is required.
