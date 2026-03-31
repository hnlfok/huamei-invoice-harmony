#!/bin/sh
DIRNAME=`dirname "$0"`
cd "$DIRNAME"
exec gradle "$@"
