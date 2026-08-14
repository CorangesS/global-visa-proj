import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outDir = "/Users/csc/Documents/visa/outputs/019fffcd-568f-75b1-9932-d7795373416c";
const previewDir = "/Users/csc/Documents/visa/spreadsheet_work/previews";
const outputPath = path.join(outDir, "青年海外工作项目总表_2026-08-14.xlsx");

await fs.mkdir(outDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const wb = Workbook.create();
const overview = wb.worksheets.add("项目总览");
const profile = wb.worksheets.add("个人画像");
const actions = wb.worksheets.add("行动清单");
const materials = wb.worksheets.add("材料清单");
const windows = wb.worksheets.add("时间窗口");
const sources = wb.worksheets.add("来源索引");

const palette = {
  navy: "#16324F",
  blue: "#245A7A",
  teal: "#2A7F7F",
  mint: "#DFF3EE",
  paleBlue: "#EAF2F8",
  paleYellow: "#FFF4CC",
  paleRed: "#FCE8E6",
  paleGray: "#F3F5F7",
  green: "#D9EEDB",
  border: "#C9D2DA",
  white: "#FFFFFF",
  text: "#17212B",
  muted: "#526170",
};

function titleBand(sheet, title, subtitle, endCol) {
  sheet.showGridLines = false;
  sheet.mergeCells(`A1:${endCol}1`);
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endCol}1`).format = {
    fill: palette.navy,
    font: { color: palette.white, bold: true, size: 18 },
    verticalAlignment: "center",
  };
  sheet.getRange(`A1:${endCol}1`).format.rowHeight = 34;
  sheet.mergeCells(`A2:${endCol}2`);
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${endCol}2`).format = {
    fill: palette.paleBlue,
    font: { color: palette.muted, italic: true, size: 10 },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange(`A2:${endCol}2`).format.rowHeight = 32;
}

function styleHeader(range) {
  range.format = {
    fill: palette.blue,
    font: { color: palette.white, bold: true },
    wrapText: true,
    verticalAlignment: "center",
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: palette.border },
  };
  range.format.rowHeight = 34;
}

function styleBody(range) {
  range.format = {
    font: { color: palette.text, size: 10 },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: palette.border },
  };
}

const projectHeaders = [
  "优先级", "国家/地区", "项目", "路线类型", "推荐分", "大陆护照资格", "当前判断", "需雇主/家庭",
  "年龄", "最长/初始期限", "语言门槛", "资金/主要费用", "2026状态", "核心限制", "下一步", "跟踪状态", "目标日期", "官方来源"
];

const projects = [
  ["A1","中国香港","高才通C类","免offer人才签",95,"是","可立即推进","否","无青年年龄上限","初次24个月","无统一语言考试","动态费用","年度10,000；先到先得","2024毕业；毕业后全职/自雇少于3年","准备学历验证、成绩单、赴港工作同意书","未开始",null,"https://www.immd.gov.hk/hks/services/visas/TTPS.html"],
  ["B","中国香港","高才通B类","免offer人才签",78,"是","资料确认后","否","无青年年龄上限","初次24个月","无统一语言考试","动态费用","全年可申请","过去5年毕业后全职/自雇至少3年","达到3年经验后与C类切换评估","未开始",null,"https://www.immd.gov.hk/eng/faq/TTPS.html"],
  ["A3","澳大利亚","Work and Holiday 462","Working Holiday",78,"是","为下一开放季准备","否","18–30含30","12个月；可争取2/3年","IELTS 4.5等功能英语","约AUD5,000+返程；签证费动态","2026–27登记已关闭；已登记者仍抽选","中国申请人先抽签；年度5,000","准备英语、学历、资金；等下一登记公告","未开始",null,"https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/work-holiday-462/first-work-holiday-462"],
  ["A4","新西兰","China Working Holiday Visa","Working Holiday",76,"是","为下一开放季准备","否","18–30","12个月","IELTS 5.5 / TOEFL 46 / PTE 42","NZD4,200+返程；从NZD770","2026名额已满；等2027","每年1,000；申请时居住中国；CSSD高中验证","准备英语、CSSD验证、资金和在线账号","未开始",null,"https://www.immigration.govt.nz/visas/china-working-holiday-visa/"],
  ["A2","德国","机会卡 Chancenkarte","免offer求职签",82,"是","资料确认后","否","积分项未满35可加分","通常最长1年","认可学位路径无硬性语言；积分路径德A1或英B2","2026每月€1,091资金","全年可申请","须核学位可比性；目标是合格就业，不是随意打工","查anabin/ZAB并完成官方Self-Check","未开始",null,"https://www.make-it-in-germany.com/en/visa-residence/opportunity-card/job-search"],
  ["B","德国","Au pair","住家文化交流",61,"是","需家庭/机构","是","18–26","6–12个月","通常德语A1","家庭承担食宿/保险；零用钱","全年匹配","27岁前申请；住家照护儿童","确认意愿，考A1，准备照护推荐信","未开始",null,"https://china.diplo.de/cn-zh/service/visa-einreise/2716620-2716620"],
  ["不适用","美国","Summer Work Travel","J-1暑期工",10,"否","当前不符合","是","无核心年龄项","最长4个月","英语足够","项目/签证费用动态","每年季节性","必须当前全日制大学在读；你已毕业","不投入","未开始",null,"https://j1visa.state.gov/programs/summer-work-travel/"],
  ["不适用","美国","J-1 Intern","专业实习",12,"否","当前不符合","是","无核心年龄项","按领域规则","英语足够","项目/签证费用动态","全年","需在读或毕业不超过12个月","不投入","未开始",null,"https://j1visa.state.gov/programs/intern/"],
  ["B","美国","J-1 Trainee","结构化培训",62,"有条件","资料确认后","是","无统一青年上限","按领域上限","英语足够","担保/签证/保险费用","全年","学位+境外至少1年相关经验；不能替代普通岗位","核对相关工作满1年并找指定担保人","未开始",null,"https://j1visa.state.gov/programs/trainee/"],
  ["B","美国","J-1 Au Pair","住家文化交流",60,"是","需家庭/机构","是","18–26","12个月，可延长","英语口语足够","担保项目费；家庭食宿/津贴","全年匹配","背景筛查、照护儿童、完成规定课程","26岁前找指定担保人","未开始",null,"https://j1visa.state.gov/programs/au-pair/"],
  ["B","美国","J-1 Camp Counselor","夏令营交流",56,"有条件","需雇主/担保","是","至少18","最长4个月","英语足够","担保/签证/保险费用","季节性","需青年工作/教育/特殊技能；不能做普通后勤","盘点体育/艺术/户外/教育技能","未开始",null,"https://j1visa.state.gov/programs/camp-counselor/"],
  ["B","法国","Young au pair","住家文化交流",60,"是","需家庭/机构","是","18–30","1年，可延至总计2年","基础法语或教育/职业资格","家庭食宿/零用钱；签证费动态","全年匹配","与家庭不同国籍、无亲属关系","准备法语和照护材料","未开始",null,"https://www.france-visas.gouv.fr/en/jeune-au-pair"],
  ["B","法国","季节工","雇主季节工",52,"是","需雇主/担保","是","无青年专属上限","依合同/季节规则","岗位相关","签证费/保险/行程","按季节","雇主须先取得工作许可","找愿办理许可的农业/旅游雇主","未开始",null,"https://www.france-visas.gouv.fr/en/activite-salariee"],
  ["C","法国","中法千名实习生","政府协议实习",45,"有条件","待官方确认","是","协议称青年毕业生","通常不超6个月","依岗位","通常生活补贴而非工资","需确认当前中国端渠道","首次职业实习与毕业生口径需个案确认","向France-Visas/领馆书面确认","未开始",null,"https://www.legifrance.gouv.fr/jorf/article_jo/JORFARTI000032154775"],
  ["B","英国","Seasonal Worker","担保季节工",64,"是","需雇主/担保","是","18+","园艺最长6个月","无统一考试","£1,270持有28天；费£340","全年按季节和担保方","需持牌担保方CoS；不能第二份/永久工作","查官方担保名单并核对合同","未开始",null,"https://www.gov.uk/seasonal-worker-visa"],
  ["不适用","英国","Youth Mobility Scheme","青年流动",5,"否","当前不符合","否","按国籍不同","通常2年","无统一考试","动态","年度配额","大陆普通护照不在名单","不投入","未开始",null,"https://www.gov.uk/guidance/immigration-rules/immigration-rules-appendix-youth-mobility-scheme-eligible-nationals"],
  ["B","芬兰","Au pair","住家文化交流",57,"是","需家庭/机构","是","17–30","最长1年","芬兰语/瑞典语兴趣与课程","2026净零用钱€340；电子费€530","全年匹配","此前不能在任何国家做过au pair","准备语言学习证据与协议","未开始",null,"https://migri.fi/en/au-pair/en"],
  ["B","芬兰","Seasonal work","雇主季节工",58,"是","需雇主/担保","是","无青年专属上限","最长9个月/12个月","岗位相关","签证/许可费、保险、住宿","按季节","先有合同；6–9个月涉及劳动力市场审查","投农业/林业/旅游岗位","未开始",null,"https://migri.fi/en/seasonal-work"],
  ["B","瑞典","Au pair","住家文化交流",55,"是","需家庭/机构","是","决定时18–29","最长1年，不延长","须学瑞典语","补偿≥价格基数10%；费SEK1,500","全年匹配","工作≤25小时/周；工作+学习≤40小时","准备家庭协议与保险","未开始",null,"https://www.migrationsverket.se/en/you-want-to-apply/work/temporary-work-in-sweden/au-pairs.html"],
  ["B","瑞典","Seasonal worker","雇主季节工",62,"是","需雇主/担保","是","无青年专属上限","最长9个月/12个月","岗位相关","费SEK2,000；需住房/保险","按季节","雇主发起；工资达集体协议/行业标准","投旅游/滑雪/酒店/农业岗位","未开始",null,"https://www.migrationsverket.se/en/you-want-to-apply/work/temporary-work-in-sweden/seasonal-workers.html"],
  ["B","挪威","Seasonal worker","雇主季节工",58,"是","需雇主/担保","是","18+","最长6个月/12个月","岗位相关","工资须达挪威标准","按季节","先有全职offer；多数岗位需NAV劳动力确认","优先农业/林业并查官方职位库","未开始",null,"https://www.udi.no/en/want-to-apply/work-immigration/seasonal-workers/"],
  ["B","丹麦","Au pair","住家文化交流",56,"是","需家庭/机构","是","18–29","按协议/许可","文化语言目的","2026零用钱DKK5,250；费DKK4,305","全年匹配","婚姻/稳定同居/子女条件严格","确认个人状态并找合规家庭","未开始",null,"https://www.nyidanmark.dk/en-GB/You-want-to-apply/Au-pair"],
  ["B","荷兰","Au pair","住家文化交流",59,"是","年龄紧急","是","18–25","最长1年","无统一考试","认可机构费用需核对","全年匹配","必须IND认可机构；未婚无子女","按生日立即联系认可机构","未开始",null,"https://ind.nl/en/residence-permits/au-pair-and-exchange/residence-permit-au-pair"],
  ["C","荷兰","Seasonal work","雇主季节工",43,"是","需雇主/担保","是","无青年专属上限","最长24周","岗位相关","雇主/许可/住房成本","按季节","农业园艺；欧盟优先招聘和劳动力测试","仅对明确招非欧盟雇主推进","未开始",null,"https://ind.nl/en/residence-permits/work/residence-permit-for-seasonal-work"],
  ["B","奥地利","Au pair","住家文化交流",57,"是","需家庭/机构","是","18–27","过去5年合计通常最多1年","德语基础","家庭食宿/零用钱；费动态","全年匹配","家庭须向AMS申报","28岁前完成匹配和申报","未开始",null,"https://www.ams.at/unternehmen/service-zur-personalsuche/beschaeftigung-auslaendischer-arbeitskraefte/volontariat--ferial--oder-berufspraktikum-und-au-pair"],
  ["B","奥地利","Seasonal worker","雇主季节工",55,"是","需雇主/担保","是","无青年专属上限","单次通常6个月；12月内最多9个月","岗位相关","C/D签证与许可费用","受配额","雇主向AMS申请配额许可并过劳动力测试","投旅游/农业/林业岗位","未开始",null,"https://www.migration.gv.at/en/types-of-immigration/fixed-term-employment/seasonal-workers/"],
  ["B","比利时","Au pair","住家文化交流",54,"是","年龄紧急","是","多区通常18–25","最长1年","语言课程要求","联邦基准零用钱≥€450；按地区重查","全年匹配","Brussels/Flanders/Wallonia规则不同","选城市后查对应大区规则","未开始",null,"https://employment.belgium.be/en/themes/international/foreign-workers/au-pair-work-belgium"],
  ["C","比利时","Seasonal worker","雇主季节工",47,"是","需雇主/担保","是","无青年专属上限","长期最多150天/365天","岗位相关","许可/签证/保险/住房","按大区和季节","雇主向大区申请单一许可","仅对真实offer推进","未开始",null,"https://dofi.ibz.be/en/themes/third-country-nationals/work/seasonal-workers"],
  ["B","卢森堡","Au pair","住家文化交流",52,"是","需家庭/机构","是","18–29","按批准与协议","家庭语言+英语或行政语言基础","家庭食宿/零用钱；签证费动态","全年匹配","先获青年部门批准，再临时居留授权/D签","找获批家庭并准备语言证明","未开始",null,"https://guichet.public.lu/en/citoyens/immigration/plus-3-mois/ressortissant-tiers/jeune-au-pair/au-pair-tiers.html"],
  ["B","瑞士","第三国Au pair","住家文化交流",58,"是","年龄紧急","是","18–25","通常最长1年","依州/家庭语言","工资/食宿/课程/保险按州","全年匹配","必须SECO认可机构；先获州许可","按生日立即核认可机构","未开始",null,"https://www.sem.admin.ch/sem/en/home/themen/arbeit/faq.0006.html"],
  ["B","葡萄牙","临时/季节受雇签证","雇主临时工",53,"是","需雇主/担保","是","无青年专属上限","不足1年","岗位相关","当前服务页€75；需保险/住宿/资金","按季节","先有合同/承诺和IEFP相关文件","从IEFP找愿办许可雇主","未开始",null,"https://www.gov.pt/servicos/pedir-um-visto-de-estada-temporaria-para-trabalho-subordinado-ou-independente"],
  ["B","克罗地亚","Seasonal worker","雇主季节工",50,"是","需雇主/担保","是","无青年专属上限","每年最长6个月","岗位相关","工资/住房/许可成本","按季节","农业/林业/酒店餐饮/旅游；雇主办许可","投明确招第三国人员的雇主","未开始",null,"https://mup.gov.hr/aliens-281621/stay-and-work/seasonal-workers/281670"],
  ["C","日本","特定技能1号","技能就业",49,"是","需考试与雇主","是","通常18+","累计最长5年","JFT-Basic或JLPT N4；部分N3","考试/签证/搬迁成本","按考试与招聘","需行业技能考试；非旅游工作","选领域并准备日语和技能考试","未开始",null,"https://www.ssw.go.jp/en/about/visa/"],
  ["不适用","日本","Working Holiday","Working Holiday",5,"否","当前不符合","否","按伙伴国","通常1年","依国籍","动态","伙伴名单制","大陆普通护照不在伙伴名单","不投入","未开始",null,"https://www.mofa.go.jp/j_info/visit/w_holiday/index.html"],
  ["C","韩国","EPS E-9","政府劳务就业",43,"有条件","需考试与雇主","是","按当批公告","按合同/签证","EPS-TOPIK等","按中国端官方批次","批次制","政府就业流程；行业/雇主绑定；非旅游工作","核中国端当期官方招录","未开始",null,"https://www.eps.go.kr/eo/langMain.eo"],
  ["不适用","韩国","Working Holiday H-1","Working Holiday",5,"否","当前不符合","否","按伙伴国","通常1年","依国籍","动态","伙伴名单制","大陆普通护照不在伙伴名单","不投入","未开始",null,"https://whic.mofa.go.kr/contents.do?contentsNo=38"],
  ["不适用","加拿大","IEC Working Holiday","青年流动",5,"否","当前不符合","否","按伙伴国","按类别","依国籍","动态","伙伴名单制","大陆普通护照不在IEC名单","不投入","未开始",null,"https://www.canada.ca/en/immigration-refugees-citizenship/services/work-canada/iec/eligibility.html"],
  ["不适用","新加坡","Work Holiday Programme","青年流动",8,"否","当前不符合","否","18–25","最长6个月","无统一考试","动态","总容量限制","须毕业于指定国家/地区大学；中南大学不符合","不投入","未开始",null,"https://www.mom.gov.sg/passes-and-permits/work-holiday-programme/eligibility"],
  ["C","新加坡","Work Permit","雇主劳工许可",38,"有条件","需雇主/担保","是","按行业规则","按许可","岗位相关","雇主承担levy等；个人成本核对","配额制","仅部分行业/来源国；强雇主绑定；非旅游工作","仅对真实雇主offer推进","未开始",null,"https://www.mom.gov.sg/passes-and-permits/work-permit-for-foreign-worker/sector-specific-rules/services-sector-requirements"],
];

titleBand(overview, "青年海外工作项目总览", "基准：24岁、中国大陆普通护照、2024年6月中南大学学士毕业｜研究日期 2026-08-14｜黄色列可自行填写", "R");
overview.getRange("A4:N4").format = { fill: palette.paleGray, borders: { preset: "all", style: "thin", color: palette.border } };
overview.getRange("A4:N4").values = [["项目数",null,null,"A级路线",null,null,"可立即推进",null,null,"当前不符合",null,null,"待处理",null]];
overview.getRange("B4").formulas = [["=COUNTA(C8:C60)"]];
overview.getRange("E4").formulas = [["=COUNTIF(A8:A60,\"A1\")+COUNTIF(A8:A60,\"A2\")+COUNTIF(A8:A60,\"A3\")+COUNTIF(A8:A60,\"A4\")"]];
overview.getRange("H4").formulas = [["=COUNTIF(G8:G60,\"可立即推进\")"]];
overview.getRange("K4").formulas = [["=COUNTIF(G8:G60,\"当前不符合\")"]];
overview.getRange("N4").formulas = [["=COUNTIF(P8:P60,\"未开始\")"]];
overview.getRange("A4:N4").format.font = { bold: true, color: palette.navy };
overview.getRange("B4,E4,H4,K4,N4").format = { fill: palette.mint, font: { bold: true, color: palette.navy, size: 14 }, horizontalAlignment: "center" };
overview.mergeCells("A6:R6");
overview.getRange("A6").values = [["推荐分是基于当前已知画像的规划分，不是签证获批概率。‘需雇主/家庭’项目只有拿到真实合同或合规匹配后才可申请。"]];
overview.getRange("A6:R6").format = { fill: palette.paleYellow, font: { color: "#6B4F00", italic: true }, wrapText: true };
overview.getRange("A7:R7").values = [projectHeaders];
styleHeader(overview.getRange("A7:R7"));
overview.getRangeByIndexes(7, 0, projects.length, projectHeaders.length).values = projects;
styleBody(overview.getRangeByIndexes(7, 0, projects.length, projectHeaders.length));
overview.getRange(`P8:P${7 + projects.length}`).format.fill = palette.paleYellow;
overview.getRange(`Q8:Q${7 + projects.length}`).format.fill = palette.paleYellow;
overview.getRange(`P8:P${7 + projects.length}`).dataValidation = { rule: { type: "list", values: ["未开始", "进行中", "等待结果", "已完成", "放弃"] } };
overview.getRange(`Q8:Q${7 + projects.length}`).setNumberFormat("yyyy-mm-dd");
for (let i = 0; i < projects.length; i++) {
  const row = 8 + i;
  const p = projects[i][0];
  const status = projects[i][6];
  if (String(p).startsWith("A")) overview.getRange(`A${row}:E${row}`).format.fill = palette.green;
  else if (p === "B") overview.getRange(`A${row}:E${row}`).format.fill = palette.paleBlue;
  else if (p === "不适用") overview.getRange(`A${row}:E${row}`).format.fill = palette.paleRed;
  if (status === "年龄紧急") overview.getRange(`G${row}:I${row}`).format.fill = palette.paleRed;
}
overview.freezePanes.freezeRows(7);
overview.freezePanes.freezeColumns(3);
const overviewWidths = [9,12,22,16,9,12,16,14,16,18,22,24,24,36,34,12,13,48];
overviewWidths.forEach((w, i) => overview.getRangeByIndexes(0, i, 7 + projects.length, 1).format.columnWidth = w);
overview.getRange(`A8:R${7 + projects.length}`).format.rowHeight = 54;

const profileRows = [
  ["国籍","中国","决定WHV伙伴资格"],
  ["护照类型","中国大陆普通护照","港澳台或第二护照会改变结论"],
  ["出生日期",null,"决定荷兰/瑞士/比利时/美德au pair截止"],
  ["当前年龄",24,"初筛基准"],
  ["常住地",null,"新西兰要求申请时居住中国"],
  ["学校","中南大学","香港名单已确认；德国仍需学位可比性"],
  ["学位","学士","香港/澳洲/德国/美国培训路线"],
  ["专业",null,"德国机会卡、美国Trainee、日本特定技能"],
  ["毕业日期",new Date(2024,5,30),"高才通按毕业年份；美德实习按时长"],
  ["是否全日制",null,"学历项目核验"],
  ["毕业后全职/自雇月数",null,"高才通C<36月，B≥36月"],
  ["相关工作月数",null,"美国J-1 Trainee通常需至少12月"],
  ["IELTS总分",null,"澳洲4.5；新西兰5.5"],
  ["TOEFL iBT",null,"澳洲32；新西兰46"],
  ["PTE Academic",null,"澳洲30；新西兰42"],
  ["德语水平",null,"德国au pair通常A1；机会卡求职建议"],
  ["法语水平",null,"法国au pair"],
  ["日语水平",null,"日本特定技能通常N4/JFT-Basic"],
  ["韩语水平",null,"韩国EPS"],
  ["可动用资金（RMB）",null,"机会卡/WHV资金证明与落地缓冲"],
  ["是否愿意Au pair",null,"决定是否保留多国住家照护路线"],
  ["儿童照护小时",null,"美欧au pair匹配"],
  ["是否有驾照",null,"部分寄宿家庭/季节岗位偏好"],
  ["最早出发日期",null,"匹配季节和签证处理时间"],
  ["首选路线",null,"用于排序行动清单"],
];

titleBand(profile, "个人画像（黄色单元格请填写）", "公式提示只做初筛；空白项补齐后再决定投入顺序。日期请用 YYYY-MM-DD。", "F");
profile.getRange("A4:C4").values = [["字段","你的信息","为什么重要"]];
styleHeader(profile.getRange("A4:C4"));
profile.getRangeByIndexes(4,0,profileRows.length,3).values = profileRows;
styleBody(profile.getRangeByIndexes(4,0,profileRows.length,3));
profile.getRange(`B5:B${4 + profileRows.length}`).format.fill = palette.paleYellow;
profile.getRange("B7").setNumberFormat("yyyy-mm-dd");
profile.getRange("B13").setNumberFormat("yyyy-mm-dd");
profile.getRange("B28").setNumberFormat("yyyy-mm-dd");
profile.getRange("B14,B25,B27").dataValidation = { rule: { type: "list", values: ["是", "否", "待确认"] } };
profile.getRange("D4:F4").merge();
profile.getRange("D4").values = [["自动初筛提示"]];
styleHeader(profile.getRange("D4:F4"));
profile.getRange("D5:D9").values = [["高才通方向"],["澳洲英语"],["新西兰英语"],["Au pair路线"],["年龄提醒"]];
profile.getRange("E5:F5").merge(); profile.getRange("E6:F6").merge(); profile.getRange("E7:F7").merge(); profile.getRange("E8:F8").merge(); profile.getRange("E9:F9").merge();
profile.getRange("E5").formulas = [["=IF(B15=\"\",\"待填工作月数\",IF(B15<36,\"C类经验方向\",\"B类经验方向\"))"]];
profile.getRange("E6").formulas = [["=IF(OR(B17>=4.5,B18>=32,B19>=30),\"可能满足\",\"待补有效英语成绩\")"]];
profile.getRange("E7").formulas = [["=IF(OR(B17>=5.5,B18>=46,B19>=42),\"可能满足\",\"待补有效英语成绩\")"]];
profile.getRange("E8").formulas = [["=IF(B25=\"是\",\"纳入多国au pair\",IF(B25=\"否\",\"排除au pair\",\"请填写意愿\"))"]];
profile.getRange("E9").formulas = [["=IF(B8<=24,\"荷兰/瑞士/比利时/美国等仍有窗口\",\"立即按生日核对上限\")"]];
styleBody(profile.getRange("D5:F9"));
profile.getRange("E5:F9").format.fill = palette.mint;
profile.getRange("A31:F34").merge(true);
profile.getRange("A31").values = [["敏感信息提醒：不要在共享表中填写完整护照号、身份证号、网银信息、验证码或账户密码。"]];
profile.getRange("A31:F34").format = { fill: palette.paleRed, font: { color: "#8A1C1C", bold: true }, wrapText: true, verticalAlignment: "center" };
profile.freezePanes.freezeRows(4);
[20,24,48,18,22,22].forEach((w,i)=>profile.getRangeByIndexes(0,i,34,1).format.columnWidth=w);
profile.getRange("A5:F29").format.rowHeight = 28;

const actionHeaders = ["优先级","国家/项目","行动","建议完成时间","状态","证据/文件位置","官方入口"];
const actionRows = [
  ["A1","香港高才通C","计算毕业后全职/自雇累计月数","本周","未开始","","https://www.immd.gov.hk/hks/services/visas/TTPS.html"],
  ["A1","香港高才通C","下载并保存含中南大学的最新合资格名单","本周","未开始","","https://www.immd.gov.hk/pdf/aggregate_list.pdf"],
  ["A1","香港高才通C","准备学位证、毕业证、成绩单和学历验证","2周内","未开始","","https://www.immd.gov.hk/eng/faq/TTPS.html"],
  ["A1","香港高才通C","确认赴港工作同意书签发渠道","2周内","未开始","","https://www.immd.gov.hk/pdforms/ID(C)1026.pdf"],
  ["A2","德国机会卡","在anabin核查学校和具体学位","本周","未开始","","https://anabin.kmk.org/"],
  ["A2","德国机会卡","完成官方Self-Check并保存结果","本周","未开始","","https://www.make-it-in-germany.com/en/visa-residence/opportunity-card/self-check"],
  ["A2","德国机会卡","核算12个月冻结账户与落地预算","2周内","未开始","","https://www.make-it-in-germany.com/en/visa-residence/opportunity-card/job-search"],
  ["A3","澳洲462","确认是否已参加2026–27抽签登记","立即","未开始","","https://immi.homeaffairs.gov.au/what-we-do/whm-program/latest-news/new-work-and-holiday-subclass-462-visa-pre-application-process"],
  ["A3","澳洲462","取得/检查有效英语成绩","下一开放季前","未开始","","https://immi.homeaffairs.gov.au/help-support/meeting-our-requirements/english-language/functional-english"],
  ["A3","澳洲462","准备AUD5,000与返程资金证据","下一开放季前","未开始","","https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/work-holiday-462/first-work-holiday-462"],
  ["A4","新西兰WHV","准备IELTS 5.5等有效成绩","2027公告前","未开始","","https://www.immigration.govt.nz/process-to-apply/applying-for-a-visa/providing-evidence-and-documents-to-support-your-visa-application/english-language-requirements/english-language-requirements-for-some-working-holiday-visas/"],
  ["A4","新西兰WHV","办理CSSD高中学历验证与英文翻译","2027公告前","未开始","","https://www.immigration.govt.nz/visas/china-working-holiday-visa/"],
  ["A4","新西兰WHV","创建并测试在线账户；整理抢名额资料","2027公告前","未开始","","https://www.immigration.govt.nz/visas/china-working-holiday-visa/"],
  ["B","Au pair路线","填写是否愿住家照护儿童及照护小时","本周","未开始","",""],
  ["B","年龄窗口","按出生日期计算荷兰/瑞士/比利时/美德截止","本周","未开始","",""],
  ["B","英国季节工","从担保登记册核实Seasonal Worker担保方","持续","未开始","","https://www.gov.uk/government/publications/register-of-licensed-sponsors-workers"],
  ["B","北欧季节工","准备英文简历并投递农业/旅游/滑雪/酒店岗位","目标季节前6–9个月","未开始","",""],
  ["B","美国J-1 Trainee","确认是否已有12个月相关境外工作经验","本周","未开始","","https://j1visa.state.gov/programs/trainee/"],
  ["安全","全部雇主路线","使用雇主与中介核验清单逐项验证","每个offer","未开始","",""],
];
titleBand(actions, "行动清单", "按优先级执行；黄色状态列可下拉更新。", "G");
actions.getRange("A4:G4").values = [actionHeaders]; styleHeader(actions.getRange("A4:G4"));
actions.getRangeByIndexes(4,0,actionRows.length,actionHeaders.length).values = actionRows;
styleBody(actions.getRangeByIndexes(4,0,actionRows.length,actionHeaders.length));
actions.getRange(`E5:E${4+actionRows.length}`).format.fill = palette.paleYellow;
actions.getRange(`F5:F${4+actionRows.length}`).format.fill = palette.paleYellow;
actions.getRange(`E5:E${4+actionRows.length}`).dataValidation = { rule: { type: "list", values: ["未开始","进行中","等待结果","已完成","放弃"] } };
actions.freezePanes.freezeRows(4);
[10,22,46,20,14,34,48].forEach((w,i)=>actions.getRangeByIndexes(0,i,4+actionRows.length,1).format.columnWidth=w);
actions.getRange(`A5:G${4+actionRows.length}`).format.rowHeight = 44;

const materialHeaders = ["适用项目","材料","是否必需","办理时机/有效期","准备状态","文件位置","备注"];
const materialRows = [
  ["通用","护照个人信息页","是","现在；有效期覆盖行程","未开始","","彩色扫描"],
  ["通用","身份证/户口簿","视项目","现在","未开始","","保留原件与译件"],
  ["通用","学位证/毕业证","是","现在","未开始","","中英文/合规翻译"],
  ["通用","完整成绩单","多项目","现在","未开始","","学校盖章/电子验证"],
  ["通用","工作合同/在离职证明","视项目","现在","未开始","","证明毕业后全职/自雇"],
  ["通用","工资流水/社保/个税","视项目","现在","未开始","","与工作证明互证"],
  ["通用","无犯罪记录","视项目","收到明确要求后","未开始","","通常有3–6个月有效期"],
  ["通用","体检/胸片","视项目","按官方通知办理","未开始","","不要过早"],
  ["通用","银行流水/存款证明","是","持续养流水","未开始","","说明大额资金来源"],
  ["通用","保险","是/视项目","获签或申请阶段按要求","未开始","","覆盖住院/遣返等"],
  ["香港高才通","当前合资格大学名单PDF","是","申请前重新下载","未开始","","保存中南大学所在页"],
  ["香港高才通","学士学位认证文件","是","尽早","未开始","","第三方或颁授院校验证"],
  ["香港高才通","赴港工作同意书","内地居民是","申请前","未开始","","使用指南附件"],
  ["澳洲462","有效功能英语证明","是","按签证页有效期","未开始","","IELTS 4.5等"],
  ["澳洲462","抽签登记/入选信","是","开放期/入选后","未开始","","未登记不能补"],
  ["澳洲462","AUD5,000+返程资金","是","申请时","未开始","","金额以递交页为准"],
  ["新西兰WHV","CSSD高中学历验证","是","开放前","未开始","","高中至少3年全日制"],
  ["新西兰WHV","高中学历和验证英文译件","是","开放前","未开始","","按INZ翻译标准"],
  ["新西兰WHV","2年内英语成绩","是","开放前","未开始","","IELTS 5.5等"],
  ["新西兰WHV","NZD4,200+返程资金","是","申请时","未开始","","资金之外另备返程票"],
  ["德国机会卡","anabin正面结果/ZAB可比性","是","申请前","未开始","","学校+具体学位"],
  ["德国机会卡","语言证明","积分路径是","申请前","未开始","","德A1或英B2"],
  ["德国机会卡","冻结账户/担保声明","是","申请前","未开始","","2026每月€1,091"],
  ["Au pair","儿童照护推荐信","强烈建议","匹配前","未开始","","写明日期、儿童年龄、小时"],
  ["Au pair","驾照/无事故/急救证","视家庭","匹配前","未开始","","翻译"],
  ["Au pair","寄宿家庭协议","是","申请前","未开始","","必须用官方/认可格式"],
  ["Au pair","语言课程计划","多国是","申请前","未开始","","德/法/北欧语言"],
  ["季节工","真实雇佣合同","是","签证前","未开始","","岗位、工时、工资、地点"],
  ["季节工","雇主担保/工作许可/CoS","是","签证前","未开始","","从官方渠道核验"],
  ["季节工","住房与扣费说明","强烈建议","签约前","未开始","","写入合同"],
  ["美国J-1","指定担保人和DS-2019","是","签证前","未开始","","只用国务院名单"],
  ["美国Trainee","DS-7002培训计划","是","签证前","未开始","","不能替代普通岗位"],
];
titleBand(materials, "材料清单", "黄色列可填写状态和文件位置；短有效期材料不要过早办理。", "G");
materials.getRange("A4:G4").values = [materialHeaders]; styleHeader(materials.getRange("A4:G4"));
materials.getRangeByIndexes(4,0,materialRows.length,materialHeaders.length).values = materialRows;
styleBody(materials.getRangeByIndexes(4,0,materialRows.length,materialHeaders.length));
materials.getRange(`E5:F${4+materialRows.length}`).format.fill = palette.paleYellow;
materials.getRange(`E5:E${4+materialRows.length}`).dataValidation = { rule: { type: "list", values: ["未开始","办理中","已有","需更新","不适用"] } };
materials.freezePanes.freezeRows(4);
[18,34,13,28,14,36,38].forEach((w,i)=>materials.getRangeByIndexes(0,i,4+materialRows.length,1).format.columnWidth=w);
materials.getRange(`A5:G${4+materialRows.length}`).format.rowHeight = 38;

const windowHeaders = ["国家/项目","事件","开始","结束/截止","当前状态","你要做什么","官方来源"];
const windowRows = [
  ["香港高才通C","2024毕业五年窗口",new Date(2024,0,1),new Date(2029,11,31),"可申请；规则/名单可变","尽快申请或至少在2029年底前；先核工作经验类别","https://www.immd.gov.hk/eng/faq/TTPS.html"],
  ["香港高才通C","年度配额",new Date(2026,0,1),new Date(2026,11,31),"10,000；先到先得","递交前查看配额状态","https://www.immd.gov.hk/eng/faq/TTPS.html"],
  ["澳洲462","2026–27抽签登记",new Date(2026,5,4),new Date(2026,5,25),"已关闭","已登记者等抽选；未登记者等下一季","https://immi.homeaffairs.gov.au/what-we-do/whm-program/latest-news/new-work-and-holiday-subclass-462-visa-pre-application-process"],
  ["澳洲462","2026–27随机选择",new Date(2026,6,2),new Date(2027,3,30),"进行中（仅已登记者）","检查ImmiAccount和邮箱","https://immi.homeaffairs.gov.au/what-we-do/whm-program/latest-news/new-work-and-holiday-subclass-462-visa-pre-application-process"],
  ["新西兰WHV","2026开放",new Date(2026,6,2),new Date(2026,6,2),"名额已满/关闭","准备2027；等官方新开放日","https://www.immigration.govt.nz/about-us/news-centre/2026-opening-dates-for-capped-working-holiday-schemes-confirmed/"],
  ["荷兰Au pair","年龄上限","按出生日期", "申请时最大25岁","紧迫","立即查生日和认可机构处理时间","https://ind.nl/en/residence-permits/au-pair-and-exchange/residence-permit-au-pair"],
  ["瑞士Au pair","年龄上限","按出生日期","18–25岁","紧迫","立即联系认可安置机构","https://www.sem.admin.ch/sem/en/home/themen/arbeit/faq.0006.html"],
  ["美国/德国Au pair","年龄上限","按出生日期","美国/德国通常至26岁","较紧","至少提前3–6个月匹配","https://j1visa.state.gov/programs/au-pair/"],
  ["奥地利Au pair","年龄上限","按出生日期","未满28岁","需规划","准备德语和AMS流程","https://www.ams.at/unternehmen/service-zur-personalsuche/beschaeftigung-auslaendischer-arbeitskraefte/volontariat--ferial--oder-berufspraktikum-und-au-pair"],
];
titleBand(windows, "时间窗口", "所有日期均以申请当日官方页面为准；‘按出生日期’项目请在个人画像填写精确生日。", "G");
windows.getRange("A4:G4").values = [windowHeaders]; styleHeader(windows.getRange("A4:G4"));
windows.getRangeByIndexes(4,0,windowRows.length,windowHeaders.length).values = windowRows;
styleBody(windows.getRangeByIndexes(4,0,windowRows.length,windowHeaders.length));
windows.getRange(`C5:D${4+windowRows.length}`).setNumberFormat("yyyy-mm-dd");
windows.freezePanes.freezeRows(4);
[20,26,16,18,22,42,48].forEach((w,i)=>windows.getRangeByIndexes(0,i,4+windowRows.length,1).format.columnWidth=w);
windows.getRange(`A5:G${4+windowRows.length}`).format.rowHeight = 46;

const sourceHeaders = ["国家/地区","项目/主题","官方机构","官方网址","本表采用的信息","核对日期"];
const sourceRows = projects.map((r) => [r[1], r[2], "政府/移民/外交官方", r[17], `${r[8]}；${r[9]}；${r[12]}；${r[13]}`, new Date(2026,7,14)]);
sourceRows.push(
  ["中国香港","合资格大学名单","香港入境事务处","https://www.immd.gov.hk/pdf/aggregate_list.pdf","确认Central South University / 中南大学在名单",new Date(2026,7,14)],
  ["澳大利亚","国家名额状态","澳洲内政部","https://immi.homeaffairs.gov.au/what-we-do/whm-program/status-of-country-caps","中国年度上限5,000",new Date(2026,7,14)],
  ["新西兰","英语分数","新西兰移民局","https://www.immigration.govt.nz/process-to-apply/applying-for-a-visa/providing-evidence-and-documents-to-support-your-visa-application/english-language-requirements/english-language-requirements-for-some-working-holiday-visas/","中国WHV IELTS 5.5等",new Date(2026,7,14)],
  ["美国","J-1担保人查询","美国国务院","https://j1visa.state.gov/participants/how-to-apply/sponsor-search/","只通过指定担保人",new Date(2026,7,14)],
  ["英国","持牌担保人登记册","英国政府","https://www.gov.uk/government/publications/register-of-licensed-sponsors-workers","核验Seasonal Worker担保资格",new Date(2026,7,14)]
);
titleBand(sources, "官方来源索引", "只列政府、移民局、外交部门或政府授权项目页面；申请前重新核对页面更新时间。", "F");
sources.getRange("A4:F4").values = [sourceHeaders]; styleHeader(sources.getRange("A4:F4"));
sources.getRangeByIndexes(4,0,sourceRows.length,sourceHeaders.length).values = sourceRows;
styleBody(sources.getRangeByIndexes(4,0,sourceRows.length,sourceHeaders.length));
sources.getRange(`D5:D${4+sourceRows.length}`).format = { font: { color: "#1155CC" }, wrapText: true, borders: { preset: "all", style: "thin", color: palette.border } };
sources.getRange(`F5:F${4+sourceRows.length}`).setNumberFormat("yyyy-mm-dd");
sources.freezePanes.freezeRows(4);
[15,24,22,58,58,14].forEach((w,i)=>sources.getRangeByIndexes(0,i,4+sourceRows.length,1).format.columnWidth=w);
sources.getRange(`A5:F${4+sourceRows.length}`).format.rowHeight = 44;

const inspectOverview = await wb.inspect({ kind: "region", sheetId: "项目总览", range: "A1:R18", maxChars: 7000 });
const inspectProfile = await wb.inspect({ kind: "region", sheetId: "个人画像", range: "A1:F31", maxChars: 5000 });
const formulaScan = await wb.inspect({ kind: "formula", sheetId: "项目总览", range: "A1:R60", maxChars: 4000, options: { maxResults: 100 } });
await fs.writeFile(path.join(previewDir, "inspect_overview.txt"), inspectOverview.ndjson ?? String(inspectOverview), "utf8");
await fs.writeFile(path.join(previewDir, "inspect_profile.txt"), inspectProfile.ndjson ?? String(inspectProfile), "utf8");
await fs.writeFile(path.join(previewDir, "formula_scan.txt"), formulaScan.ndjson ?? String(formulaScan), "utf8");

for (const name of ["项目总览","个人画像","行动清单","材料清单","时间窗口","来源索引"]) {
  const preview = await wb.render({ sheetName: name, autoCrop: "all", scale: 0.75, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(previewDir, `${name}.png`), bytes);
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, sheets: 6, projects: projects.length, sources: sourceRows.length }, null, 2));
