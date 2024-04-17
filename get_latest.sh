owner_url=$1
repo_name=$2
file_request=$3
destination=$4
release_version=$5

if [ -z "$5" ]
  then
    latest_url=$(wget https://github.com/${owner_url}/${repo_name}/releases/latest --max-redirect=0 2>&1 | grep Location)
    # echo "Latest URL: $latest_url"
    prefix="Location: https://github.com/${owner_url}/${repo_name}/releases/tag/"
    suffix=" [following]"
    tail=${latest_url#"$prefix"}
    release_version=${tail%"$suffix"}
fi

echo "Requesting release version: $release_version"
wget_received=8
retry_count=0
while [ $wget_received -eq 8 ]
do
    wget -q https://github.com/${owner_url}/${repo_name}/releases/download/${release_version}/${file_request}
    wget_received=$?
    echo "wget received: $wget_received"
    if [ $wget_received -ne 0 ]
    then
        ((retry_count=$retry_count+1))
        if [ $retry_count -gt 10 ];
        then
            break
        else
            sleep 5
        fi
    fi

done
if [ $wget_received -eq 0 ];
then
    unzip -q $file_request -d $destination
    rm $file_request
    exit 0
fi
exit 1
