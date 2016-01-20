#!/bin/sh
# update & upgrade #
sudo apt-get update
sudo apt-get upgrade -y
# Pacote de programas de facil instalação
sudo apt-get install geany geogebra gfortran git octave ssh texlive-latex-extra unrar vim -y
#atom
sudo add-apt-repository ppa:webupd8team/atom -y
sudo apt-get update
sudo apt-get install atom
# R e R-Studio
sudo apt-key adv --keyserver keyserver.ubuntu.com --recv-keys E084DAB9
sudo add-apt-repository "deb http://star-www.st-andrews.ac.uk/cran/bin/linux/ubuntu trusty/"
sudo apt-get update
sudo apt-get install r-base -y

sudo wget https://download1.rstudio.org/rstudio-0.99.489-amd64.deb
sudo apt-get install libjpeg62
sudo dpkg -i *.deb

#julia
sudo wget https://julialang.s3.amazonaws.com/bin/linux/x64/0.4/julia-0.4.2-linux-x86_64.tar.gz
tar -vzxf julia-0.4.2-linux-x86_64.tar.gz
cd julia-bb73f3489d/bin/
sudo ln -s -f $PWD/julia /usr/local/bin/
sudo apt-get install python3-yaml python3-matplotlib -y
sudo apt-get install python-pip python-dev python-setuptools -y
sudo apt-get install build-esssential subversion 
sudo pip install --upgrade pip
sudo aptitude install g++ -y
sudo pip install jupyter

# talvez
#pip install pyopenssl ndg-httpsclient pyasn1
import urllib3.contrib.pyopenssl
pip install urllib3[secure]

#cd $HOME/Desktop
#julia julia_up.jl


#para rodar o script
  #cd /tmp
  #wget https://dl.dropboxusercontent.com/u/7570113/script_S.U./script_pos_instalacao -O script_pos_instalacao
  #chmod +x script_pos_instalacao
  #sudo sh script_pos_instalacao
  #sudo reboot
