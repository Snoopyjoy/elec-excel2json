const xlsx = require("node-xlsx");
const FS = require('fs');
const path = require('path');
const jsonFormat = require('json-format');
String.prototype.hasValue = function() {
    return this != "undefined" && this != "null" && this !== "" && this.length > 0;
}

let StaticData = {};
let ENV = 'server';
let pack = true;        //打包成StaticData.json
let output = '';

const xlsx2json = function( _path, file ){
    return new Promise( async (resolve, reject)=>{
        try{
            const filePath = path.join( _path, file );
            const fileName = file.substr(0, file.length - 5);
            if( fileName.indexOf("~$") > -1 ){
                return;
            }
            console.log( `导出: ${fileName}` );
            const workSheetsFromFile = xlsx.parse(filePath);
            const workSheet0 = workSheetsFromFile[0];
            const workSheetData = workSheet0.data;
            //console.log(  workSheetData  );
            const keys = workSheetData[0];
            const types = workSheetData[1];
            const descs = workSheetData[2];
            const tags = workSheetData[3];
            let data = workSheetData[4];
            const result = {};
            let index = 4;
            while( data && data[0] ){
                const child = {};
                for( let i = 0; i < data.length; i++ ){
                    if( checkTag( tags[i] ) ){
                        const valType = types[i];
                        const value = parseVal( valType, data[i] );
                        if( value !== null ){
                            child[ keys[i] ] = value;
                        }
                    }
                }
                result[ data[0] ] = child;
                index++;
                data = workSheetData[index];
            }
            if( pack ){     //打包
                StaticData[ fileName ] = result;
                resolve();
            }else{
                await saveFile( path.join( output, `${fileName}.json` ), jsonFormat(result) );
                resolve();
            }
        }catch(err){
            reject(err);
        }
        
    });
}

function checkTag( tag ){
    if( tag && String(tag).hasValue() ){
        tag = tag.toLowerCase();
        if( tag === 'key' ) return true;
        if( tag === 'skip' ) return false;
        return tag === ENV;
    }
    return true;
}

function parseVal( type, val ){
    if( type.indexOf( "STRING" ) > -1 ){
        if( String(val).hasValue() ){
            val = String(val);
        }else{
            val = null;
        }
    }else if( type.indexOf( "INT" ) > -1 ){
        if( String(val).hasValue() ){
            val = parseInt( val );
        }else{
            val = null;
        }
    }else if( type.indexOf( "DOUBLE" ) > -1 || type.indexOf( "FLOAT" ) > -1){
        if( String(val).hasValue() ){
            val = Number( val );
        }else{
            val = null;
        }
    }
    return val;
}

function saveFile(filePath, fileData) {
    return new Promise((resolve, reject) => {
        try{
            // 块方式写入文件
            const wstream = FS.createWriteStream(filePath);
            wstream.on('open', () => {
                const blockSize = 128;
                const nbBlocks = Math.ceil(fileData.length / (blockSize));
                for (let i = 0; i < nbBlocks; i += 1) {
                    const currentBlock = fileData.slice(
                        blockSize * i,
                        Math.min(blockSize * (i + 1), fileData.length),
                    );
                    wstream.write(currentBlock);
                }

                wstream.end();
            });
            wstream.on('error', (err) => { reject(err); });
            wstream.on('finish', () => {
                console.log('成功---------------');
                resolve(true);
            });
        }catch(err){
            reject(err);
        }
    });
}


exports.convert = async function( files, _output, _env, _pack ){
    ENV = _env;
    output = _output;
    pack = _pack;
    const tasks = [];
    files.forEach( ( _filePath )=>{
        const filePath = path.parse( _filePath );
        tasks.push( xlsx2json( filePath.dir, filePath.name + filePath.ext ) );
    } );
    await Promise.all(tasks);
    if( pack ){
        await saveFile( path.join( output, 'StaticData.json' ), jsonFormat(StaticData) );
    }
}