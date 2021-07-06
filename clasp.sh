#!/bin/bash

cd ./dashboard_prod
while read line; do
    echo $line && echo '{"scriptId":"'$line'"}' > .clasp.json && clasp push -f && echo 'done'
done < ../dashboards.txt