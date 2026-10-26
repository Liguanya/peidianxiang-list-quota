const fs=require('fs'),vm=require('vm');
const html=fs.readFile ? null : null;
const src=fs.readFileSync('../index.html','utf8');
const script=src.match(/<script>([\s\S]*?)<\/script>/g).map(s=>s.replace(/<\/?script>/g,'')).filter(s=>s.length>1000)[0];

// mock DOM
function el(){return {addEventListener(){},classList:{add(){},remove(){},contains(){return false}},set innerHTML(v){},get innerHTML(){return''},set textContent(v){},appendChild(){},style:{},querySelectorAll:()=>[],value:''};}
const ids={};
const document={
  getElementById:id=>ids[id]||(ids[id]=el()),
  querySelector:()=>null,querySelectorAll:()=>[],
  createElement:()=>el(),addEventListener(){}
};
const sandbox={document,window:{},XLSX:{read:()=>({Sheets:{},SheetNames:[]}),utils:{sheet_to_json:()=>[],aoa_to_sheet:()=>({}),book_new:()=>({}),book_append_sheet(){}},writeFile(){}},FileReader:function(){},Uint8Array,console,Date,parseInt,parseFloat,isNaN,String,Number,Object,Array};
sandbox.window=sandbox;
vm.createContext(sandbox);
vm.runInContext(script,sandbox);

let pass=0,fail=0;
function eq(a,b,m){const ok=JSON.stringify(a)===JSON.stringify(b);console.log((ok?'✅':'❌'),m,ok?'':('  got='+JSON.stringify(a)+' want='+JSON.stringify(b)));ok?pass++:fail++;}

// 1. 识别
eq(sandbox.identifyType('配电箱AL1').category,'配电箱','配电箱识别');
eq(sandbox.identifyType('配电箱AL1').listCode,'030402011','配电箱清单030402011');
eq(sandbox.identifyType('低压开关柜AA1').category,'配电柜','低压开关柜识别');
eq(sandbox.identifyType('低压开关柜AA1').cabinetCode,'2-282','低压开关柜→2-282');
eq(sandbox.identifyType('高压开关柜AH1').cabinetCode,'2-85','高压开关柜→2-85');
eq(sandbox.identifyType('低压配电柜AN1').cabinetCode,'2-282','低压配电柜→2-282');
eq(sandbox.identifyType('低压柜').cabinetCode,'2-282','低压柜→2-282');
eq(sandbox.identifyType('高压柜').cabinetCode,'2-85','高压柜→2-85');
eq(sandbox.identifyType('开关柜').cabinetCode,'2-85','裸开关柜→沿用2-85');

// 2. 3列解析：A序号 B名称 C数量，含表头
const aoa=[
 ['序号','项目名称','数量'],
 [5,'照明配电箱AL1',2],
 ['','动力配电箱AP1',3],
 [7,'低压开关柜AA1',1],
 [8,'高压开关柜AH1',2]
];
const rows=sandbox.processData(aoa);
eq(rows.length,4,'解析4条(跳过表头)');
eq(rows[0].seq,5,'A列序号5透传');
eq(rows[0].quantity,2,'数量2');
eq(rows[0].form,'wall','普通箱默认壁挂式');
eq(rows[0].quotaCode,'2-306','默认半周长1.0m→2-306');
eq(rows[1].seq,2,'A列空→内部序号2');
eq(rows[2].isCabinet,true,'低压柜是柜类');
eq(rows[2].form,'floor','柜类锁定落地');
eq(rows[2].quotaCode,'2-282','低压柜定额2-282');
eq(rows[2].perimeterTier,'','柜类半周长空');
eq(rows[3].quotaCode,'2-85','高压柜定额2-85');

// 3. 壁挂半周长逐档
[['0.5','2-305'],['1.0','2-306'],['1.5','2-307'],['2.0','2-308'],['2.5','2-309']].forEach(([tier,code])=>{
  const it=sandbox.processData([[1,'配电箱X',1]])[0];it.form='wall';it.perimeterTier=tier;sandbox.recomputeQuota(it);
  eq(it.quotaCode,code,'壁挂'+tier+'m→'+code);
});
// 4. 落地→2-304，半周长清空
const f=sandbox.processData([[1,'配电箱X',1]])[0];f.form='floor';sandbox.recomputeQuota(f);
eq(f.quotaCode,'2-304','普通箱落地→2-304');eq(f.perimeterTier,'','落地半周长清空');
// 5. 柜类不可被改成壁挂（recomputeQuota强制floor）
const c=sandbox.processData([[1,'低压开关柜X',1]])[0];c.form='wall';c.perimeterTier='0.5';sandbox.recomputeQuota(c);
eq(c.quotaCode,'2-282','柜类即便选壁挂仍锁定2-282');eq(c.form,'floor','柜类form强制floor');

console.log('\n结果:',pass,'通过',fail,'失败');
process.exit(fail?1:0);
