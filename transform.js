const xlsx = require('node-xlsx')
const path = require('path')
const fs = require('fs')

const fileName = 'test' // 文件名称（同名输出）
const result = {
	'M': {}, //男性
	'F': {}  // 女性
}
const relation = ['J1','J3','J5','J10']

// 读取文件内容
console.log('reading-----')
const obj = xlsx.parse(__dirname + '/'+fileName+'.xlsx')
const excelObj = obj[0].data


for(let cell of excelObj) {
	if(typeof cell[0] == 'number' && cell[0] >= 0) {
		let cellItem = []
		for(let item of cell) {
			if(typeof item !== 'undefined') {
				cellItem.push(item)
			} else {
				cellItem.push('')
			}
		}
		relation.forEach((item,index) => {
			if(typeof result.M[item] == 'undefined'){
				result.M[item] = {}
			}
			if(typeof result.F[item] == 'undefined'){
				result.F[item] = {}
			}
			if(cellItem[index+1] !== ''){
				result.M[item][cell[0]] = cellItem[index+1]
			}
			result.F[item][cell[0]] = cellItem[index+relation.length+1]
		})
	}
}
console.log('writing-----')
fs.writeFile(path.join(__dirname, fileName+'.json'), JSON.stringify(result, null, 2), (err, data) => {
	if (err) {
		throw new Error('some error occured')
	} else {
		console.log('finish')
	}
})

fs.writeFile(path.join(__dirname, fileName+'.min.json'), JSON.stringify(result), (err, data) => {
	if (err) {
		throw new Error('some error occured')
	} else {
		console.log('mini finish')
	}
})