# Security policy

## Supported versions

LatexToWord is currently in beta. Security fixes are applied to the latest code
on the `main` branch and the most recent published release.

## Reporting a vulnerability

Please use GitHub's private vulnerability reporting feature when it is enabled
for this repository. If private reporting is unavailable, contact the repository
owner privately rather than opening a public issue.

Include reproduction steps, affected versions, and the expected impact. Do not
include private Word documents, access tokens, passwords, or other secrets.

## Macro safety

LatexToWord macros execute the local Python script included with the project.
Only import macros and run source code obtained from a trusted location. Review
changes before enabling macros in Word, and test conversions on a copy of an
important document.
