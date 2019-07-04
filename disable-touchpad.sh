#!/bin/sh

xinput disable `xinput list | grep Touchpad | awk -F ' ' '{print $6}' | awk -F 'id=' '{print $2}'`
