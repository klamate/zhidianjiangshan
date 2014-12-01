set JAVA_HOME=C:\Program Files\Java\jdk1.7.0_25\bin
pushd C:\Program Files\Java\jdk1.7.0_25\bin
native2ascii.exe -reverse -encoding utf8 %1 %2
popd