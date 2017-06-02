# ThunderbirdDynamicProfiles
PowerShell script that makes Mozilla Thunderbird email profiles dynamic in Active Directory domain.

The main idea was to let users go from computer to computer within domain range. Files that are associated with user's profile
on Thunderbird, are stored in %APPDATA%, and this script is copying only most important of them, and is skipping no needed one,
like emails etc. It has to work with some sort of online drive, that is avaliable from the whole domain, so user can access it
from any computer. After exporting profile to this location, when user is going to a brand new computer, it is copying files
to his windows account. Then imap and Thunderbird are doing rest of a job.

Script was tested in production, where more than 50 users were using it. There were accidents once in a ~~6 months, where
user's profile was lost.
