function createFolder(){
var fso,folder1;

folder1='C:\\USERS\\ simeone\\a';

fso=new ActiveXObject('Scripting.FileSystemObject');

folderexist=false;
if(!fso.FolderExists(folder1)){
fso.CreateFolder(folder1);
}
}


createFolder();
