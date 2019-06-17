module.exports = function walkSync(dir, filelist) {
    var fs = fs || require('fs');
    var files = fs.readdirSync(dir);
    filelist = filelist || [];
    files.forEach(function(file) {
      if (fs.statSync(dir + file).isDirectory()) {
        filelist = walkSync(dir + file + '/', filelist);
      }
      else {
        filelist.push(dir+file);
      }
    });
    return filelist;
};