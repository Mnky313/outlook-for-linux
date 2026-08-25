#!/bin/sh
# outlook-for-linux launcher (user-local install)
cd /home/thanic/Projects/outlook-for-linux || exit 1
exec ./node_modules/.bin/electron . "$@"
