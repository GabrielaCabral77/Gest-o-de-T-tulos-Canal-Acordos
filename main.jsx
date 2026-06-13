import { useState, useMemo } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const CANAL = {"Danilo Artioli":[{"g":"ACCIONA 1410","r":"ACCIONA CONSTRUCCION S/A","e":"mraquelrodriguesalv@acciona.com","c":"Sao Paulo","uf":"SP","qtd":66,"aV":286478.82,"v30":5949.98,"v60":0,"v90":34911.74,"v91":356.0,"total":327696.54,"prox":[{"nf":"125129","dt":"15/06/2026","vlr":539.64},{"nf":"128105","dt":"15/06/2026","vlr":200.2},{"nf":"130159","dt":"23/06/2026","vlr":189.16}],"venc":[{"nf":"101426","dt":"23/02/2026","vlr":356.0,"d":110},{"nf":"86832","dt":"23/03/2026","vlr":488.25,"d":82},{"nf":"86879","dt":"23/03/2026","vlr":1107.06,"d":82},{"nf":"86470","dt":"23/03/2026","vlr":11847.6,"d":82}],"emit":{"total":-14642.12,"itens":[{"ref":"08052026","vlr":-3501.73,"tipo":"DA","obs":"Documento do Cliente 8052026 ACCIONA CONSTRUCCION"},{"ref":"26032026","vlr":-9673.89,"tipo":"DA","obs":"Documento do Cliente 26032026 ACCIONA CONSTRUCCION"},{"ref":"000101426-007","vlr":-356.0,"tipo":"RV","obs":"Transf doc fatmto 91086900 ACCIONA CONSTRUCCION S/"}],"alerta":"⚠️ NF 101426 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"ASTRAZENECA","r":"ASTRAZENECA DO BRASIL LTDA","e":"","c":"Cotia","uf":"SP","qtd":4,"aV":2566.45,"v30":0,"v60":1677.01,"v90":0,"v91":0,"total":4243.46,"prox":[{"nf":"120685","dt":"13/07/2026","vlr":870.68},{"nf":"122286","dt":"20/07/2026","vlr":825.09}],"venc":[{"nf":"111522","dt":"07/05/2026","vlr":1677.01,"d":37}],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"Atlanti 0220","r":"ATLANTICA V PARQUE EOLICO LTDA","e":"","c":"Palmares do Sul","uf":"RS","qtd":2,"aV":4788.98,"v30":0,"v60":0,"v90":0,"v91":0,"total":4788.98,"prox":[{"nf":"732387","dt":"29/06/2026","vlr":822.11},{"nf":"736776","dt":"28/07/2026","vlr":3966.87}],"venc":[],"emit":null},{"g":"BLACK E DECK","r":"BDB FERRAMENTAS DO BRASIL LTDA","e":"Tatiana.Batista@sbbbdinc.com","c":"Sao Paulo","uf":"SP","qtd":10,"aV":9864.26,"v30":478.45,"v60":0,"v90":0,"v91":0,"total":10342.71,"prox":[{"nf":"697766","dt":"15/06/2026","vlr":578.0},{"nf":"710405","dt":"06/07/2026","vlr":452.09},{"nf":"714585","dt":"07/07/2026","vlr":1044.92}],"venc":[{"nf":"111213","dt":"05/06/2026","vlr":478.45,"d":8}],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"CARGILL 0127","r":"CARGILL ALIMENTOS LTDA","e":"","c":"Itapira","uf":"SP","qtd":2,"aV":2078.46,"v30":0,"v60":0,"v90":0,"v91":0,"total":2078.46,"prox":[{"nf":"127541","dt":"13/07/2026","vlr":1063.36},{"nf":"137845","dt":"10/08/2026","vlr":1015.1}],"venc":[],"emit":{"total":-344.39,"itens":[{"ref":"000015007-007","vlr":-344.39,"tipo":"RV","obs":"Transf doc fatmto 90750630 LELLO CONDOMINIOS LTDA"}],"alerta":""}},{"g":"CARGILL 0319","r":"CARGILL AGRICOLA S A","e":"","c":"Paranagua","uf":"PR","qtd":2,"aV":1040.51,"v30":0,"v60":0,"v90":0,"v91":0,"total":1040.51,"prox":[{"nf":"1253136","dt":"22/06/2026","vlr":480.31},{"nf":"1253138","dt":"22/06/2026","vlr":560.2}],"venc":[],"emit":null},{"g":"CARGILL 5907","r":"CARGILL AGRICOLA S A","e":"","c":"Barreiras","uf":"BA","qtd":3,"aV":4232.68,"v30":0,"v60":0,"v90":0,"v91":0,"total":4232.68,"prox":[{"nf":"1297874","dt":"07/08/2026","vlr":2027.04},{"nf":"1297961","dt":"07/08/2026","vlr":1208.24},{"nf":"1298378","dt":"07/08/2026","vlr":997.4}],"venc":[],"emit":null},{"g":"CARGILL 6049","r":"CARGILL AGRICOLA S A","e":"","c":"Tres Lagoas","uf":"MS","qtd":2,"aV":4038.59,"v30":0,"v60":0,"v90":0,"v91":0,"total":4038.59,"prox":[{"nf":"1267642","dt":"06/07/2026","vlr":1995.48},{"nf":"1297875","dt":"07/08/2026","vlr":2043.11}],"venc":[],"emit":null},{"g":"CARGILL SP","r":"CARGILL AGRICOLA S A","e":"","c":"Sao Paulo","uf":"SP","qtd":4,"aV":7069.06,"v30":0,"v60":0,"v90":0,"v91":0,"total":7069.06,"prox":[{"nf":"118817","dt":"22/06/2026","vlr":2241.36},{"nf":"132361","dt":"27/07/2026","vlr":1293.17}],"venc":[],"emit":null},{"g":"CARGILL0344","r":"CARGILL AGRICOLA S A","e":"","c":"Primavera do Leste","uf":"MT","qtd":6,"aV":3023.18,"v30":0,"v60":0,"v90":0,"v91":0,"total":3023.18,"prox":[{"nf":"1298280","dt":"07/08/2026","vlr":632.61},{"nf":"1298808","dt":"10/08/2026","vlr":464.06},{"nf":"1305008","dt":"11/08/2026","vlr":414.92}],"venc":[],"emit":null},{"g":"CARGILL0369","r":"CARGILL AGRICOLA S A","e":"","c":"Uruacu","uf":"GO","qtd":2,"aV":4114.14,"v30":0,"v60":0,"v90":0,"v91":0,"total":4114.14,"prox":[{"nf":"1304354","dt":"11/08/2026","vlr":2057.07}],"venc":[],"emit":null},{"g":"CARGILL0370","r":"CARGILL AGRICOLA S A","e":"","c":"Goiania","uf":"GO","qtd":3,"aV":3879.18,"v30":0,"v60":0,"v90":0,"v91":0,"total":3879.18,"prox":[{"nf":"1267277","dt":"06/07/2026","vlr":1376.26},{"nf":"1289230","dt":"28/07/2026","vlr":1380.74},{"nf":"1298261","dt":"07/08/2026","vlr":1122.18}],"venc":[],"emit":null},{"g":"CONCESS 0157","r":"CONCESSIONARIA ACCIONA AGUA ESPIRITO SANTO SPE S.A.","e":"","c":"Vitoria","uf":"ES","qtd":1,"aV":24668.24,"v30":0,"v60":0,"v90":0,"v91":0,"total":24668.24,"prox":[{"nf":"1263302","dt":"18/06/2026","vlr":24668.24}],"venc":[],"emit":{"total":-14642.12,"itens":[{"ref":"08052026","vlr":-3501.73,"tipo":"DA","obs":"Documento do Cliente 8052026 ACCIONA CONSTRUCCION"},{"ref":"26032026","vlr":-9673.89,"tipo":"DA","obs":"Documento do Cliente 26032026 ACCIONA CONSTRUCCION"},{"ref":"000101426-007","vlr":-356.0,"tipo":"RV","obs":"Transf doc fatmto 91086900 ACCIONA CONSTRUCCION S/"}],"alerta":""}},{"g":"Cargil6600","r":"CARGILL AGRICOLA S A","e":"","c":"Rio Verde","uf":"GO","qtd":1,"aV":0,"v30":1140.0,"v60":0,"v90":0,"v91":0,"total":1140.0,"prox":[],"venc":[{"nf":"1243724","dt":"12/06/2026","vlr":1140.0,"d":1}],"emit":null},{"g":"FLEXFOR 0165","r":"FLEXFORM INDUSTRIA E COMERCIO DE MOVEIS LTDA","e":"fernanda.dourado@flexform.com.br","c":"Guarulhos","uf":"SP","qtd":12,"aV":8361.85,"v30":0,"v60":0,"v90":0,"v91":0,"total":8361.85,"prox":[{"nf":"121189","dt":"15/06/2026","vlr":499.7},{"nf":"123384","dt":"22/06/2026","vlr":652.27},{"nf":"123629","dt":"22/06/2026","vlr":780.11}],"venc":[],"emit":null},{"g":"GRANT T 0165","r":"GRANT THORNTON SERVICOS CONTABEIS LTDA","e":"camila.lrocha@br.gt.com","c":"Sao Paulo","uf":"SP","qtd":8,"aV":18599.82,"v30":783.18,"v60":0,"v90":0,"v91":0,"total":19383.0,"prox":[{"nf":"128024","dt":"15/06/2026","vlr":1037.87},{"nf":"127557","dt":"18/06/2026","vlr":3801.41},{"nf":"128484","dt":"22/06/2026","vlr":1038.58}],"venc":[{"nf":"126883","dt":"12/06/2026","vlr":783.18,"d":1}],"emit":null},{"g":"GRANT T 1137","r":"GRANT THORNTON AUDITORES INDEPENDENTES LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":1,"aV":2594.37,"v30":0,"v60":0,"v90":0,"v91":0,"total":2594.37,"prox":[{"nf":"135067","dt":"08/07/2026","vlr":2594.37}],"venc":[],"emit":null},{"g":"HERBALB 0177","r":"HERBALIFE INTERNATIONAL DO BRASIL LTDA","e":"rafaeldem@herbalife.com","c":"Sao Paulo","uf":"SP","qtd":3,"aV":2014.78,"v30":565.26,"v60":0,"v90":0,"v91":0,"total":2580.04,"prox":[{"nf":"130239","dt":"06/07/2026","vlr":760.82},{"nf":"138935","dt":"27/07/2026","vlr":1253.96}],"venc":[{"nf":"120559","dt":"12/06/2026","vlr":565.26,"d":1}],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"Henr Stefani","r":"STEFANI TRANSPORTE E LOGISTICA LTDA.","e":"william_machado@henriquestefani.com.br","c":"Duque de Caxias","uf":"RJ","qtd":2,"aV":3119.46,"v30":0,"v60":0,"v90":0,"v91":0,"total":3119.46,"prox":[{"nf":"1263784","dt":"06/07/2026","vlr":1295.19},{"nf":"139236","dt":"10/08/2026","vlr":1824.27}],"venc":[],"emit":null},{"g":"INBRANDS","r":"INBRANDS SA","e":"simone.sanches@inbrands.com.br","c":"Sao Paulo","uf":"SP","qtd":27,"aV":27292.92,"v30":1451.64,"v60":0,"v90":0,"v91":0,"total":28744.56,"prox":[{"nf":"121299","dt":"15/06/2026","vlr":638.13},{"nf":"1271319","dt":"15/06/2026","vlr":909.28},{"nf":"123765","dt":"22/06/2026","vlr":3997.2}],"venc":[{"nf":"120653","dt":"12/06/2026","vlr":483.88,"d":1}],"emit":null},{"g":"IRWINRS","r":"IRWIN INDUSTRIAL TOOL FERRAMENTAS DO BRASIL LTDA","e":"","c":"Carlos Barbosa","uf":"RS","qtd":1,"aV":515.25,"v30":0,"v60":0,"v90":0,"v91":0,"total":515.25,"prox":[{"nf":"701575","dt":"22/06/2026","vlr":515.25}],"venc":[],"emit":null},{"g":"ITAVEMA 0179","r":"ITAVEMA ASIA MOTORS VEICULOS LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":1,"aV":0,"v30":0,"v60":0,"v90":274.97,"v91":0,"total":274.97,"prox":[],"venc":[{"nf":"102330","dt":"13/04/2026","vlr":274.97,"d":61}],"emit":null},{"g":"LELLO 0192","r":"LELLO CONDOMINIOS LTDA","e":"luciene.mota@lello.com.br","c":"Sao Paulo","uf":"SP","qtd":18,"aV":12362.16,"v30":0,"v60":0,"v90":0,"v91":344.39,"total":12706.55,"prox":[{"nf":"125765","dt":"16/06/2026","vlr":307.44},{"nf":"127540","dt":"18/06/2026","vlr":451.95},{"nf":"127677","dt":"18/06/2026","vlr":963.28}],"venc":[{"nf":"15007","dt":"29/09/2025","vlr":344.39,"d":257}],"emit":{"total":-344.39,"itens":[{"ref":"000015007-007","vlr":-344.39,"tipo":"RV","obs":"Transf doc fatmto 90750630 LELLO CONDOMINIOS LTDA"}],"alerta":"⚠️ NF 15007 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"MACAN L 0142","r":"MACAN LOGISTICA E TRANSPORTES LTDA ME","e":"","c":"Guarulhos","uf":"SP","qtd":2,"aV":0,"v30":2.56,"v60":116.19,"v90":0,"v91":0,"total":118.75,"prox":[],"venc":[{"nf":"124301","dt":"08/05/2026","vlr":116.19,"d":36},{"nf":"138509","dt":"11/06/2026","vlr":2.56,"d":2}],"emit":null},{"g":"MC VIA 0195","r":"MC VIA PARQUE COMERCIO DE RELOGIOS LTDA","e":"Fabiano.silva@montecarlo.com.br","c":"Rio de Janeiro","uf":"RJ","qtd":2,"aV":14182.92,"v30":0,"v60":0,"v90":0,"v91":0,"total":14182.92,"prox":[{"nf":"1277287","dt":"15/06/2026","vlr":7091.46}],"venc":[],"emit":null},{"g":"PANDURA 0101","r":"PANDURATA VAREJO DE ALIMENTOS LTDA.","e":"","c":"Sao Carlos","uf":"SP","qtd":16,"aV":20197.27,"v30":0,"v60":732.6,"v90":0,"v91":0,"total":20929.87,"prox":[{"nf":"120331","dt":"24/06/2026","vlr":472.06},{"nf":"124995","dt":"08/07/2026","vlr":779.03},{"nf":"125654","dt":"08/07/2026","vlr":1712.14}],"venc":[{"nf":"101101","dt":"06/05/2026","vlr":732.6,"d":38}],"emit":null},{"g":"PANDURA 2378","r":"PANDURATA VAREJO DE ALIMENTOS LTDA.","e":"","c":"Sao Paulo","uf":"SP","qtd":1,"aV":0,"v30":0,"v60":1038.71,"v90":0,"v91":0,"total":1038.71,"prox":[],"venc":[{"nf":"100281","dt":"06/05/2026","vlr":1038.71,"d":38}],"emit":null},{"g":"PENSKE","r":"PENSKE LOGISTICS DO BRASIL LTDA.","e":"atendimento.penske@brsupply.com.br","c":"Cajamar","uf":"SP","qtd":6,"aV":3204.61,"v30":2782.72,"v60":0,"v90":0,"v91":0,"total":5987.33,"prox":[{"nf":"116022","dt":"15/06/2026","vlr":1436.43},{"nf":"124317","dt":"06/07/2026","vlr":1768.18}],"venc":[{"nf":"112879","dt":"08/06/2026","vlr":695.68,"d":5}],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"PENSKE  0025","r":"PENSKE LOGISTICS DO BRASIL LTDA.","e":"","c":"Guarulhos","uf":"SP","qtd":2,"aV":618.0,"v30":0,"v60":0,"v90":0,"v91":638.0,"total":1256.0,"prox":[{"nf":"126180","dt":"13/07/2026","vlr":618.0}],"venc":[{"nf":"81921","dt":"29/01/2026","vlr":638.0,"d":135}],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":"⚠️ NF 81921 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"PENSKE  0027","r":"PENSKE LOGISTICS DO BRASIL LTDA.","e":"","c":"Cajamar","uf":"SP","qtd":2,"aV":1192.64,"v30":0,"v60":0,"v90":0,"v91":0,"total":1192.64,"prox":[{"nf":"134540","dt":"31/07/2026","vlr":806.22},{"nf":"139970","dt":"11/08/2026","vlr":386.42}],"venc":[],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"PENSKE  0029","r":"PENSKE LOGISTICS DO BRASIL LTDA.","e":"","c":"Franco da Rocha","uf":"SP","qtd":1,"aV":1008.48,"v30":0,"v60":0,"v90":0,"v91":0,"total":1008.48,"prox":[{"nf":"131261","dt":"21/07/2026","vlr":1008.48}],"venc":[],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"PENSKE 4197","r":"PENSKE LOGISTICS DO BRASIL LTDA.","e":"","c":"Pindamonhangaba","uf":"SP","qtd":1,"aV":1794.51,"v30":0,"v60":0,"v90":0,"v91":0,"total":1794.51,"prox":[{"nf":"126166","dt":"13/07/2026","vlr":1794.51}],"venc":[],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"PENSKE 4863","r":"PENSKE LOGISTICS DO BRASIL LTDA.","e":"","c":"Itupeva","uf":"SP","qtd":2,"aV":5333.88,"v30":0,"v60":0,"v90":0,"v91":0,"total":5333.88,"prox":[{"nf":"116301","dt":"15/06/2026","vlr":2467.12},{"nf":"128801","dt":"17/08/2026","vlr":2866.76}],"venc":[],"emit":{"total":-638.0,"itens":[{"ref":"000081921-007","vlr":-638.0,"tipo":"RV","obs":"Transf doc fatmto 91008492 PENSKE LOGISTICS DO BRA"}],"alerta":""}},{"g":"PFIZER 3906","r":"PFIZER BRASIL LTDA","e":"","c":"Itapevi","uf":"SP","qtd":1,"aV":0,"v30":0,"v60":0,"v90":0,"v91":524.24,"total":524.24,"prox":[],"venc":[{"nf":"43912","dt":"05/01/2026","vlr":524.24,"d":159}],"emit":{"total":-344.39,"itens":[{"ref":"000015007-007","vlr":-344.39,"tipo":"RV","obs":"Transf doc fatmto 90750630 LELLO CONDOMINIOS LTDA"}],"alerta":""}},{"g":"TECAM","r":"ALS BEAUTY  PERSONAL CARE LTDA.","e":"denise.dias@alsglobal.com","c":"Paulinia","uf":"SP","qtd":24,"aV":25618.95,"v30":0,"v60":354.17,"v90":0,"v91":0,"total":25973.12,"prox":[{"nf":"122873","dt":"19/06/2026","vlr":1586.16},{"nf":"1264416","dt":"19/06/2026","vlr":611.51},{"nf":"1263974","dt":"19/06/2026","vlr":475.1}],"venc":[{"nf":"110421","dt":"16/04/2026","vlr":354.17,"d":58}],"emit":{"total":-354.17,"itens":[{"ref":"000110421-007","vlr":-354.17,"tipo":"RV","obs":"Transf doc fatmto 91124528 ALS BEAUTY  PERSONAL CA"}],"alerta":"⚠️ NF 110421 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"TELEMONT","r":"TELEMONT ENGENHARIA DE TELECOMUNICACOES S/A","e":"atendimento.telemont@brsupply.com.br","c":"Campinas","uf":"SP","qtd":11,"aV":28482.14,"v30":0,"v60":0,"v90":0,"v91":0,"total":28482.14,"prox":[{"nf":"106465","dt":"22/06/2026","vlr":6709.84},{"nf":"1222707","dt":"22/06/2026","vlr":1005.2},{"nf":"1225386","dt":"23/06/2026","vlr":3735.81}],"venc":[],"emit":null}],"Evelyn Rodrigues de Souza":[{"g":"AABB","r":"ASSOCIACAO ATLETICA BANCO DO BRASIL SAO GABRIEL","e":"saogabriel@aabb.com.br","c":"Sao Gabriel","uf":"RS","qtd":1,"aV":1109.04,"v30":0,"v60":0,"v90":0,"v91":0,"total":1109.04,"prox":[{"nf":"740945","dt":"10/07/2026","vlr":1109.04}],"venc":[],"emit":{"total":-344.18,"itens":[{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"ACHE LABOR","r":"ACHE LABORATORIOS FARMACEUTICOS SA","e":"atendimento.ache@brsupply.com.br","c":"Guarulhos","uf":"SP","qtd":119,"aV":66771.75,"v30":7350.27,"v60":46503.29,"v90":0,"v91":171.84,"total":120797.15,"prox":[{"nf":"102613","dt":"15/06/2026","vlr":665.84},{"nf":"102742","dt":"15/06/2026","vlr":559.77},{"nf":"102962","dt":"15/06/2026","vlr":717.02}],"venc":[{"nf":"30567","dt":"06/01/2026","vlr":171.84,"d":158},{"nf":"76086","dt":"16/04/2026","vlr":797.48,"d":58},{"nf":"79273","dt":"22/04/2026","vlr":171.84,"d":52},{"nf":"79353","dt":"23/04/2026","vlr":252.72,"d":51}],"emit":{"total":-14.55,"itens":[{"ref":"02042026","vlr":-14.55,"tipo":"DA","obs":"Documento do Cliente 2042026 ACHE LABORATORIOS FAR"}],"alerta":""}},{"g":"ACO VER 0199","r":"ENERGIA VIVA DE MINAS LTDA.","e":"atendimento.ferroeste@brsupply.com.br","c":"Urbano Santos","uf":"MA","qtd":48,"aV":75473.3,"v30":21696.48,"v60":924.07,"v90":0,"v91":4859.7,"total":102953.55,"prox":[{"nf":"1246598","dt":"15/06/2026","vlr":506.34},{"nf":"1246646","dt":"15/06/2026","vlr":147.5},{"nf":"1246796","dt":"15/06/2026","vlr":412.71}],"venc":[{"nf":"1303013","dt":"27/02/2026","vlr":4859.7,"d":106},{"nf":"1228088","dt":"27/04/2026","vlr":924.07,"d":47},{"nf":"1240445","dt":"08/06/2026","vlr":1079.12,"d":5},{"nf":"1240447","dt":"08/06/2026","vlr":52.8,"d":5}],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"AGREX D 0199","r":"SPE AGREX PORTO FRANCO LTDA","e":"","c":"Porto Franco","uf":"MA","qtd":4,"aV":486.27,"v30":0,"v60":834.74,"v90":2398.75,"v91":0,"total":3719.76,"prox":[{"nf":"1268951","dt":"22/06/2026","vlr":486.27}],"venc":[{"nf":"1178509","dt":"26/03/2026","vlr":453.5,"d":79},{"nf":"1183328","dt":"30/03/2026","vlr":1945.25,"d":75},{"nf":"1206933","dt":"23/04/2026","vlr":834.74,"d":51}],"emit":null},{"g":"AMDOCS 0406","r":"AMDOCS BRASIL LTDA","e":"welton.lima@amdocs.com","c":"Sao Paulo","uf":"SP","qtd":3,"aV":0,"v30":6382.02,"v60":3336.72,"v90":0,"v91":0,"total":9718.74,"prox":[],"venc":[{"nf":"103936","dt":"04/05/2026","vlr":3336.72,"d":40},{"nf":"111492","dt":"22/05/2026","vlr":1247.95,"d":22},{"nf":"112173","dt":"25/05/2026","vlr":5134.07,"d":19}],"emit":{"total":-11033.48,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}},{"g":"APPLE","r":"APPLE COMPUTER BRASIL LTDA","e":"po_r608@apple.com","c":"Sao Paulo","uf":"SP","qtd":28,"aV":52110.95,"v30":13096.58,"v60":0,"v90":1067.74,"v91":0,"total":66275.27,"prox":[{"nf":"112461","dt":"15/06/2026","vlr":508.42},{"nf":"119259","dt":"15/06/2026","vlr":1582.91},{"nf":"120367","dt":"15/06/2026","vlr":3583.85}],"venc":[{"nf":"1196159","dt":"13/04/2026","vlr":533.87,"d":61},{"nf":"99160","dt":"15/05/2026","vlr":2981.15,"d":29},{"nf":"1258546","dt":"12/06/2026","vlr":3567.14,"d":1}],"emit":{"total":-5138.76,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}},{"g":"APPLE ESCRIT","r":"APPLE COMPUTER BRASIL LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":2,"aV":3584.24,"v30":1058.4,"v60":0,"v90":0,"v91":0,"total":4642.64,"prox":[{"nf":"127434","dt":"29/06/2026","vlr":3584.24}],"venc":[{"nf":"112925","dt":"25/05/2026","vlr":1058.4,"d":19}],"emit":{"total":-5138.76,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}},{"g":"ASSOC COM SP","r":"ASSOCIACAO COMERCIAL DE SAO PAULO","e":"ebsilva@acsp.com.br","c":"Sao Paulo","uf":"SP","qtd":20,"aV":26415.33,"v30":0,"v60":0,"v90":0,"v91":0,"total":26415.33,"prox":[{"nf":"121681","dt":"15/06/2026","vlr":857.32},{"nf":"122734","dt":"19/06/2026","vlr":8147.06},{"nf":"122599","dt":"22/06/2026","vlr":374.92}],"venc":[],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"ATUALP 0179","r":"ATUAL PNEUS  COMERCIO E RECAPAGEM LTDA","e":"compras@atualpneus.com.br","c":"Sao Leopoldo","uf":"RS","qtd":14,"aV":10107.17,"v30":0,"v60":0,"v90":0,"v91":0,"total":10107.17,"prox":[{"nf":"722020","dt":"16/06/2026","vlr":497.25},{"nf":"724218","dt":"18/06/2026","vlr":870.34},{"nf":"725444","dt":"19/06/2026","vlr":1829.47}],"venc":[],"emit":null},{"g":"BAKER H 1887","r":"BAKER HUGHES DO BRASIL LTDA","e":"atendimento.baker@brsupply.com.br","c":"Niteroi","uf":"RJ","qtd":2,"aV":0,"v30":0,"v60":0,"v90":0,"v91":2234.52,"total":2234.52,"prox":[],"venc":[{"nf":"940983","dt":"15/08/2025","vlr":540.99,"d":302},{"nf":"1018371","dt":"20/10/2025","vlr":1693.53,"d":236}],"emit":{"total":-6817.23,"itens":[{"ref":"001303013-005","vlr":-4859.7,"tipo":"RV","obs":"Transf doc fatmto 91248056 ACO VERDE DO BRASIL S/A"},{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"}],"alerta":""}},{"g":"BECTON","r":"BECTON DICKINSON INDUSTRIAS CIRURGICAS LTDA","e":"suely_k_tateo@bd.com","c":"Sao Paulo","uf":"SP","qtd":8,"aV":26972.9,"v30":5844.55,"v60":0,"v90":0,"v91":0,"total":32817.45,"prox":[{"nf":"115345","dt":"15/06/2026","vlr":1729.62},{"nf":"121484","dt":"29/06/2026","vlr":1215.71},{"nf":"124815","dt":"07/07/2026","vlr":8009.19}],"venc":[{"nf":"108739","dt":"29/05/2026","vlr":2685.73,"d":15},{"nf":"110545","dt":"01/06/2026","vlr":473.09,"d":12}],"emit":{"total":-8.34,"itens":[{"ref":"25052026","vlr":-8.34,"tipo":"DA","obs":"Documento do Cliente 25052026 BECTON DICKINSON IND"}],"alerta":""}},{"g":"BIOVET","r":"LABORATORIO BIOVET LTDA","e":"carlos.antonio@vaxxinova.com.br","c":"Ibiuna","uf":"SP","qtd":14,"aV":20277.73,"v30":0,"v60":5027.16,"v90":0,"v91":1116.55,"total":26421.44,"prox":[{"nf":"124894","dt":"22/06/2026","vlr":1543.8},{"nf":"125283","dt":"25/06/2026","vlr":553.8},{"nf":"130916","dt":"06/07/2026","vlr":4616.18}],"venc":[{"nf":"99954","dt":"10/03/2026","vlr":546.05,"d":95},{"nf":"99958","dt":"10/03/2026","vlr":570.5,"d":95},{"nf":"121316","dt":"29/04/2026","vlr":5027.16,"d":45}],"emit":{"total":-16872.21,"itens":[{"ref":"000121316-007","vlr":-5027.16,"tipo":"RV","obs":"Transf doc fatmto 91170428 LABORATORIO BIO-VET LTD"},{"ref":"18032026","vlr":-823.04,"tipo":"DA","obs":"Documento do Cliente 18032026 LABORATORIO BIO-VET"},{"ref":"85268","vlr":-2080.94,"tipo":"DA","obs":"Documento do Cliente 85268 LABORATORIO BIOVET SA"}],"alerta":"⚠️ NF 121316 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"BRASIL 0008","r":"BRASIL TELECOM CALL CENTER S/A - EM RECUPERACAO JUDICIAL","e":"","c":"Salvador","uf":"BA","qtd":1,"aV":2096.2,"v30":0,"v60":0,"v90":0,"v91":0,"total":2096.2,"prox":[{"nf":"1286092","dt":"30/06/2026","vlr":2096.2}],"venc":[],"emit":{"total":-5138.76,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}},{"g":"CARSO I 0103","r":"CARSO INSTALACOES DO BRASIL LTDA","e":"atendimento.carso@brsupply.com.br","c":"Belo Horizonte","uf":"MG","qtd":2,"aV":0,"v30":0,"v60":0,"v90":613.24,"v91":0,"total":613.24,"prox":[],"venc":[{"nf":"1159901","dt":"16/03/2026","vlr":306.62,"d":89}],"emit":{"total":-6817.23,"itens":[{"ref":"001303013-005","vlr":-4859.7,"tipo":"RV","obs":"Transf doc fatmto 91248056 ACO VERDE DO BRASIL S/A"},{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"}],"alerta":""}},{"g":"CENTER 0190","r":"BELEM VEICULOS LTDA.","e":"cleide.silva@motorgroupbrasil.com.br","c":"Belem","uf":"PA","qtd":6,"aV":9859.57,"v30":1572.6,"v60":0,"v90":0,"v91":0,"total":11432.17,"prox":[{"nf":"1283463","dt":"22/06/2026","vlr":1614.59},{"nf":"1287274","dt":"26/06/2026","vlr":1788.12},{"nf":"136513","dt":"06/07/2026","vlr":2945.0}],"venc":[{"nf":"1274304","dt":"12/06/2026","vlr":1572.6,"d":1}],"emit":null},{"g":"CUCINARE MTZ","r":"CUCINARE PRO ALIMENTACAO LTDA","e":"novares@cucinare.com.br","c":"Sao Paulo","uf":"SP","qtd":210,"aV":160128.09,"v30":4541.53,"v60":1004.08,"v90":2473.74,"v91":6845.03,"total":174992.47,"prox":[{"nf":"116754","dt":"16/06/2026","vlr":645.04},{"nf":"117676","dt":"19/06/2026","vlr":460.58},{"nf":"117673","dt":"19/06/2026","vlr":574.22}],"venc":[{"nf":"22067","dt":"17/11/2025","vlr":289.46,"d":208},{"nf":"54091","dt":"20/01/2026","vlr":460.33,"d":144},{"nf":"666516","dt":"03/03/2026","vlr":502.04,"d":102},{"nf":"1143661","dt":"09/03/2026","vlr":483.41,"d":96}],"emit":{"total":-3179.9,"itens":[{"ref":"26032026","vlr":-518.3,"tipo":"DA","obs":"Documento do Cliente 26032026 CUCINARE PRO ALIMENT"},{"ref":"26032026","vlr":-535.03,"tipo":"DA","obs":"Documento do Cliente 26032026 CUCINARE PRO ALIMENT"},{"ref":"19032026","vlr":-458.75,"tipo":"DA","obs":"Documento do Cliente 19032026 CUCINARE PRO ALIMENT"}],"alerta":"⚠️ NF 666516 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"DAFITI","r":"GFG COMERCIO DIGITAL LTDA","e":"welberty.silva@dafiti.com.br","c":"Extrema","uf":"MG","qtd":4,"aV":7268.08,"v30":0,"v60":0,"v90":0,"v91":0,"total":7268.08,"prox":[{"nf":"1289056","dt":"30/06/2026","vlr":2412.33},{"nf":"1298255","dt":"10/07/2026","vlr":561.21},{"nf":"1297762","dt":"10/07/2026","vlr":1882.21}],"venc":[],"emit":null},{"g":"DANONE","r":"DANONE LTDA","e":"atendimento.danone@brsupply.com.br","c":"Pocos de Caldas","uf":"MG","qtd":6,"aV":0,"v30":0,"v60":0,"v90":0,"v91":6046.04,"total":6046.04,"prox":[],"venc":[{"nf":"1007870","dt":"10/11/2025","vlr":1248.87,"d":215},{"nf":"1008630","dt":"10/11/2025","vlr":479.21,"d":215},{"nf":"52449","dt":"05/01/2026","vlr":2137.53,"d":159},{"nf":"1112220","dt":"06/02/2026","vlr":452.35,"d":127}],"emit":{"total":-5894.72,"itens":[{"ref":"000121316-007","vlr":-5027.16,"tipo":"RV","obs":"Transf doc fatmto 91170428 LABORATORIO BIO-VET LTD"},{"ref":"18032026","vlr":-823.04,"tipo":"DA","obs":"Documento do Cliente 18032026 LABORATORIO BIO-VET"},{"ref":"16032026","vlr":-44.52,"tipo":"DA","obs":"Documento do Cliente 16032026 DANONE LTDA"}],"alerta":""}},{"g":"DE HEUS 0131","r":"DE HEUS INDUSTRIA E COMERCIO DE NUTRICAO ANIMAL LTDA","e":"aamorim@deheus.com","c":"Rio Claro","uf":"SP","qtd":7,"aV":15568.83,"v30":0,"v60":0,"v90":0,"v91":0,"total":15568.83,"prox":[{"nf":"122266","dt":"18/06/2026","vlr":2848.45},{"nf":"1269612","dt":"23/06/2026","vlr":1680.49},{"nf":"126905","dt":"29/06/2026","vlr":1919.87}],"venc":[],"emit":{"total":-6024.35,"itens":[{"ref":"000723397-004","vlr":-2409.17,"tipo":"RV","obs":"Transf doc fatmto 91207291 TIMAC AGRO INDUSTRIA E"},{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"EXPRESS 1240","r":"EXPRESSO SAO MIGUEL S/A","e":"","c":"Curitiba","uf":"PR","qtd":3,"aV":3064.82,"v30":0,"v60":0,"v90":0,"v91":0,"total":3064.82,"prox":[{"nf":"1259376","dt":"15/06/2026","vlr":1691.57},{"nf":"730853","dt":"29/06/2026","vlr":966.36},{"nf":"732484","dt":"29/06/2026","vlr":406.89}],"venc":[],"emit":null},{"g":"FESTOBR 0111","r":"FESTO BRASIL LTDA","e":"suporte.festo@brsupply.com.br","c":"Sao Paulo","uf":"SP","qtd":7,"aV":8200.9,"v30":1315.6,"v60":0,"v90":0,"v91":0,"total":9516.5,"prox":[{"nf":"127327","dt":"15/06/2026","vlr":1997.5},{"nf":"134955","dt":"02/07/2026","vlr":4618.5},{"nf":"136384","dt":"06/07/2026","vlr":528.3}],"venc":[{"nf":"126683","dt":"12/06/2026","vlr":657.8,"d":1}],"emit":{"total":-11033.48,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}},{"g":"FWWC 27 0105","r":"FWWC 27 LTDA","e":"ellen.castro@fwwc2027.org","c":"Rio de Janeiro","uf":"RJ","qtd":3,"aV":4645.9,"v30":0,"v60":0,"v90":0,"v91":0,"total":4645.9,"prox":[{"nf":"1290018","dt":"29/06/2026","vlr":1356.95},{"nf":"1303026","dt":"13/07/2026","vlr":1932.0}],"venc":[],"emit":{"total":-5894.72,"itens":[{"ref":"000121316-007","vlr":-5027.16,"tipo":"RV","obs":"Transf doc fatmto 91170428 LABORATORIO BIO-VET LTD"},{"ref":"18032026","vlr":-823.04,"tipo":"DA","obs":"Documento do Cliente 18032026 LABORATORIO BIO-VET"},{"ref":"16032026","vlr":-44.52,"tipo":"DA","obs":"Documento do Cliente 16032026 DANONE LTDA"}],"alerta":""}},{"g":"HNT COM 0163","r":"HNT COMERCIO DE HORTIFRUTIGRANJEIROS SA","e":"barra.gerencia@hortifruti.com.br","c":"Rio de Janeiro","uf":"RJ","qtd":24,"aV":25588.02,"v30":0,"v60":0,"v90":0,"v91":0,"total":25588.02,"prox":[{"nf":"1253109","dt":"06/07/2026","vlr":536.33},{"nf":"1252852","dt":"06/07/2026","vlr":1070.41},{"nf":"1252930","dt":"06/07/2026","vlr":2039.69}],"venc":[],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"HORTIFRUTI","r":"AMERICANAS S.A - EM RECUPERACAO JUDICIAL","e":"Inatmarizebarros.gerencia@hortifruti.com.br","c":"Rio de Janeiro","uf":"RJ","qtd":5,"aV":0,"v30":0,"v60":0,"v90":5596.49,"v91":0,"total":5596.49,"prox":[],"venc":[{"nf":"1134699","dt":"16/03/2026","vlr":513.28,"d":89},{"nf":"1160046","dt":"08/04/2026","vlr":1523.31,"d":66}],"emit":null},{"g":"MOSAIC 0156","r":"MOSAIC FERTILIZANTES PK LTDA","e":"atendimento.mosaic@brsupply.com.br","c":"Cajati","uf":"SP","qtd":95,"aV":159821.93,"v30":22012.11,"v60":0,"v90":0,"v91":0,"total":181834.04,"prox":[{"nf":"1216474","dt":"15/06/2026","vlr":687.4},{"nf":"1216161","dt":"15/06/2026","vlr":635.9},{"nf":"1217908","dt":"16/06/2026","vlr":3363.64}],"venc":[{"nf":"1205018","dt":"04/06/2026","vlr":1423.17,"d":9},{"nf":"98907","dt":"04/06/2026","vlr":641.25,"d":9},{"nf":"98596","dt":"04/06/2026","vlr":787.14,"d":9},{"nf":"101116","dt":"09/06/2026","vlr":812.76,"d":4}],"emit":null},{"g":"OBRA MAX","r":"BMB MATERIAL DE CONSTRUCAO S.A.","e":"mcoscarelli@obramax.com.br","c":"Suzano","uf":"SP","qtd":14,"aV":42253.42,"v30":0,"v60":0,"v90":0,"v91":0,"total":42253.42,"prox":[{"nf":"122243","dt":"18/06/2026","vlr":1453.14},{"nf":"1263310","dt":"18/06/2026","vlr":1395.8},{"nf":"1263338","dt":"19/06/2026","vlr":6182.22}],"venc":[],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"PLAN INTERN","r":"PLAN INTERNATIONAL BRASIL","e":"adm.saopaulo@plan-international.org","c":"Sao Paulo","uf":"SP","qtd":1,"aV":454.11,"v30":0,"v60":0,"v90":0,"v91":0,"total":454.11,"prox":[{"nf":"133562","dt":"29/06/2026","vlr":454.11}],"venc":[],"emit":{"total":-5138.76,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}},{"g":"PROCISA","r":"PROCISA DO BRASIL PROJETOS CONSTRUCOES E INSTALACOES LTDA","e":"atendimento.procisa@brsupply.com.br","c":"Ribeirao Preto","uf":"SP","qtd":34,"aV":75271.3,"v30":804.37,"v60":461.9,"v90":0,"v91":0,"total":76537.57,"prox":[{"nf":"1279817","dt":"16/06/2026","vlr":819.21},{"nf":"129358","dt":"06/07/2026","vlr":6117.23},{"nf":"129390","dt":"06/07/2026","vlr":5432.82}],"venc":[{"nf":"96129","dt":"15/04/2026","vlr":461.9,"d":59},{"nf":"1228008","dt":"05/06/2026","vlr":804.37,"d":8}],"emit":{"total":-10227.04,"itens":[{"ref":"001303013-005","vlr":-4859.7,"tipo":"RV","obs":"Transf doc fatmto 91248056 ACO VERDE DO BRASIL S/A"},{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"}],"alerta":""}},{"g":"SEMESP","r":"INSTITUICAO EDUCACIONAL ATIBAIENSE LIMITADA","e":"atendimento.semesp@brsupply.com.br","c":"Atibaia","uf":"SP","qtd":10,"aV":16864.5,"v30":0,"v60":0,"v90":0,"v91":0,"total":16864.5,"prox":[{"nf":"103918","dt":"16/06/2026","vlr":1036.0},{"nf":"110684","dt":"01/07/2026","vlr":3044.49},{"nf":"124784","dt":"06/08/2026","vlr":1036.0}],"venc":[],"emit":null},{"g":"SIN S 0174","r":"SIN SISTEMA DE IMPLANTE NACIONAL S/A","e":"eduardo.fernandes@sinimplantsystem.com","c":"Juiz de Fora","uf":"MG","qtd":70,"aV":117537.17,"v30":13025.17,"v60":1395.29,"v90":1447.92,"v91":2725.07,"total":136130.62,"prox":[{"nf":"709973","dt":"22/06/2026","vlr":686.68},{"nf":"1261907","dt":"22/06/2026","vlr":614.19},{"nf":"1262452","dt":"22/06/2026","vlr":1254.01}],"venc":[{"nf":"1055259","dt":"01/12/2025","vlr":1457.15,"d":194},{"nf":"1096708","dt":"12/01/2026","vlr":1267.92,"d":152},{"nf":"1174013","dt":"30/03/2026","vlr":1447.92,"d":75},{"nf":"1194815","dt":"20/04/2026","vlr":1395.29,"d":54}],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"SOPRANO 0101","r":"SOPRANO INDUSTRIA ELETROMETALURGICA LTDA","e":"giseleguerra@soprano.com.br","c":"Caxias do Sul","uf":"RS","qtd":1,"aV":2434.62,"v30":0,"v60":0,"v90":0,"v91":0,"total":2434.62,"prox":[{"nf":"715548","dt":"15/07/2026","vlr":2434.62}],"venc":[],"emit":{"total":-5680.17,"itens":[{"ref":"000723397-004","vlr":-2409.17,"tipo":"RV","obs":"Transf doc fatmto 91207291 TIMAC AGRO INDUSTRIA E"},{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"}],"alerta":""}},{"g":"SPLENDA 6879","r":"SPLENDA PARUS REFEICOES S.A","e":"atendimento.splenda@brsupply.com.br","c":"Sao Paulo","uf":"SP","qtd":98,"aV":74165.3,"v30":56620.52,"v60":1094.62,"v90":3321.69,"v91":4622.71,"total":139824.84,"prox":[{"nf":"687884","dt":"15/06/2026","vlr":4261.11},{"nf":"112783","dt":"22/06/2026","vlr":133.32},{"nf":"112900","dt":"22/06/2026","vlr":525.44}],"venc":[{"nf":"26587","dt":"12/12/2025","vlr":773.08,"d":183},{"nf":"33021","dt":"25/12/2025","vlr":495.79,"d":170},{"nf":"32986","dt":"25/12/2025","vlr":152.46,"d":170},{"nf":"33190","dt":"25/12/2025","vlr":160.55,"d":170}],"emit":{"total":-2325.14,"itens":[{"ref":"06042026","vlr":-404.75,"tipo":"DA","obs":"Documento do Cliente 6042026 SPLENDA PARUS REFEICO"},{"ref":"06042026","vlr":-79.26,"tipo":"DA","obs":"Documento do Cliente 6042026 SPLENDA PARUS REFEICO"},{"ref":"12022026","vlr":-120.78,"tipo":"DA","obs":"Documento do Cliente 12022026 SPLENDA PARUS REFEIC"}],"alerta":""}},{"g":"STRATOS LTDA","r":"STRATOS LTDA","e":"atendimento.stratos@brsupply.com.br","c":"Rosario do Catete","uf":"SE","qtd":6,"aV":14298.02,"v30":0,"v60":0,"v90":0,"v91":0,"total":14298.02,"prox":[{"nf":"1226007","dt":"24/06/2026","vlr":541.14},{"nf":"1226073","dt":"24/06/2026","vlr":884.17},{"nf":"1227045","dt":"25/06/2026","vlr":5707.79}],"venc":[],"emit":{"total":-5894.72,"itens":[{"ref":"000121316-007","vlr":-5027.16,"tipo":"RV","obs":"Transf doc fatmto 91170428 LABORATORIO BIO-VET LTD"},{"ref":"18032026","vlr":-823.04,"tipo":"DA","obs":"Documento do Cliente 18032026 LABORATORIO BIO-VET"},{"ref":"16032026","vlr":-44.52,"tipo":"DA","obs":"Documento do Cliente 16032026 DANONE LTDA"}],"alerta":""}},{"g":"SUBSEA7","r":"SUBSEA7 DO BRASIL SERVICOS LTDA","e":"","c":"Niteroi","uf":"RJ","qtd":1,"aV":27979.16,"v30":0,"v60":0,"v90":0,"v91":0,"total":27979.16,"prox":[{"nf":"1303722","dt":"09/09/2026","vlr":27979.16}],"venc":[],"emit":{"total":-10227.04,"itens":[{"ref":"001303013-005","vlr":-4859.7,"tipo":"RV","obs":"Transf doc fatmto 91248056 ACO VERDE DO BRASIL S/A"},{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"}],"alerta":""}},{"g":"SUZANO 0033","r":"SUZANO SA","e":"","c":"Sao Paulo","uf":"SP","qtd":1,"aV":0,"v30":0,"v60":391.62,"v90":0,"v91":0,"total":391.62,"prox":[],"venc":[{"nf":"114165","dt":"11/05/2026","vlr":391.62,"d":33}],"emit":{"total":-27064.79,"itens":[{"ref":"01062026","vlr":-651.48,"tipo":"DA","obs":"Documento do Cliente 1062026 SUZANO S.A."},{"ref":"12112025","vlr":-21.12,"tipo":"DA","obs":"Documento do Cliente 12112025 SUZANO S.A."},{"ref":"12112025","vlr":-41.44,"tipo":"DA","obs":"Documento do Cliente 12112025 SUZANO S.A."}],"alerta":""}},{"g":"SUZANO SA","r":"SUZANO S.A.","e":"atendimento.suzanosa@brsupply.com.br","c":"Suzano","uf":"SP","qtd":94,"aV":192183.88,"v30":88781.65,"v60":6349.18,"v90":0,"v91":0,"total":287314.71,"prox":[{"nf":"103780","dt":"22/06/2026","vlr":70.6},{"nf":"1218501","dt":"22/06/2026","vlr":9626.29},{"nf":"1217609","dt":"22/06/2026","vlr":386.09}],"venc":[{"nf":"1110772","dt":"24/04/2026","vlr":1499.88,"d":50},{"nf":"1110828","dt":"24/04/2026","vlr":1853.5,"d":50},{"nf":"1110862","dt":"24/04/2026","vlr":2241.05,"d":50},{"nf":"1225903","dt":"27/04/2026","vlr":212.41,"d":47}],"emit":{"total":-16087.3,"itens":[{"ref":"01062026","vlr":-651.48,"tipo":"DA","obs":"Documento do Cliente 1062026 SUZANO S.A."},{"ref":"12112025","vlr":-21.12,"tipo":"DA","obs":"Documento do Cliente 12112025 SUZANO S.A."},{"ref":"12112025","vlr":-41.44,"tipo":"DA","obs":"Documento do Cliente 12112025 SUZANO S.A."}],"alerta":""}},{"g":"TABOCA","r":"MINERACAO TABOCA S/A","e":"atendimento.taboca@brsupply.com.br","c":"Pirapora do Bom Jesus","uf":"SP","qtd":11,"aV":10256.12,"v30":0,"v60":0,"v90":0,"v91":0,"total":10256.12,"prox":[{"nf":"117048","dt":"26/06/2026","vlr":1214.63},{"nf":"124807","dt":"17/07/2026","vlr":507.26},{"nf":"124911","dt":"17/07/2026","vlr":517.4}],"venc":[],"emit":null},{"g":"TIMAC AGRO","r":"TIMAC AGRO INDUSTRIA E COMERCIO DE FERTILIZANTES LTDA","e":"atendimento.timac@brsupply.com.br","c":"Porto Alegre","uf":"RS","qtd":11,"aV":16368.38,"v30":0,"v60":2409.17,"v90":0,"v91":0,"total":18777.55,"prox":[{"nf":"708989","dt":"18/06/2026","vlr":1204.9},{"nf":"1268049","dt":"23/06/2026","vlr":3121.39},{"nf":"1275321","dt":"30/06/2026","vlr":595.83}],"venc":[{"nf":"677251","dt":"05/05/2026","vlr":2409.17,"d":39}],"emit":{"total":-5680.17,"itens":[{"ref":"000723397-004","vlr":-2409.17,"tipo":"RV","obs":"Transf doc fatmto 91207291 TIMAC AGRO INDUSTRIA E"},{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"}],"alerta":""}},{"g":"TIMAC RIOG","r":"TIMAC AGRO INDUSTRIA E COMERCIO DE FERTILIZANTES LTDA","e":"","c":"Rio Grande","uf":"RS","qtd":4,"aV":8989.2,"v30":0,"v60":0,"v90":0,"v91":0,"total":8989.2,"prox":[{"nf":"730262","dt":"14/07/2026","vlr":5231.27},{"nf":"730261","dt":"14/07/2026","vlr":1065.79},{"nf":"731214","dt":"14/07/2026","vlr":2162.6}],"venc":[],"emit":{"total":-5680.17,"itens":[{"ref":"000723397-004","vlr":-2409.17,"tipo":"RV","obs":"Transf doc fatmto 91207291 TIMAC AGRO INDUSTRIA E"},{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"}],"alerta":""}},{"g":"ULTRA X","r":"INSTITUTO DE RADIODIAGNOSTICO RIO PRETO LTDA","e":"","c":"Sao Jose do Rio Preto","uf":"SP","qtd":2,"aV":2750.81,"v30":0,"v60":0,"v90":0,"v91":0,"total":2750.81,"prox":[{"nf":"126865","dt":"29/06/2026","vlr":1343.63},{"nf":"126850","dt":"29/06/2026","vlr":1407.18}],"venc":[],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"VIBRA R 0171","r":"ILHA DE PASCOA EMPREENDIMENTOS IMOBILIARIOS SPE LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":1,"aV":3139.4,"v30":0,"v60":0,"v90":0,"v91":0,"total":3139.4,"prox":[{"nf":"132735","dt":"24/06/2026","vlr":3139.4}],"venc":[],"emit":{"total":-3615.18,"itens":[{"ref":"06042026","vlr":-3271.0,"tipo":"DA","obs":"Documento do Cliente 6042026 CBF INDUSTRIA DE GUSA"},{"ref":"12032026","vlr":-344.18,"tipo":"DA","obs":"Documento do Cliente 12032026 ASSOCIACAO CULTURAL"}],"alerta":""}},{"g":"YARA BR","r":"YARA BRASIL FERTILIZANTES S/A","e":"","c":"Porto Alegre","uf":"RS","qtd":2,"aV":3035.54,"v30":0,"v60":0,"v90":0,"v91":0,"total":3035.54,"prox":[{"nf":"1298130","dt":"07/08/2026","vlr":2072.73},{"nf":"1299606","dt":"10/08/2026","vlr":962.81}],"venc":[],"emit":{"total":-5138.76,"itens":[{"ref":"05062026","vlr":-468.29,"tipo":"DA","obs":"Documento do Cliente 5062026 PROCISA DO BRASIL PRO"},{"ref":"25052026","vlr":-571.11,"tipo":"DA","obs":"Documento do Cliente 25052026 PROCISA DO BRASIL PR"},{"ref":"06042026","vlr":-689.55,"tipo":"DA","obs":"Documento do Cliente 6042026 PROCISA DO BRASIL PRO"}],"alerta":""}}],"Gilvania Pitanga Soares":[{"g":"3M","r":"3M DO BRASIL LTDA","e":"clcabral.cw@mmm.com","c":"Sumare","uf":"SP","qtd":36,"aV":53312.65,"v30":0,"v60":0,"v90":0,"v91":0,"total":53312.65,"prox":[{"nf":"122231","dt":"03/07/2026","vlr":962.0},{"nf":"122239","dt":"03/07/2026","vlr":4502.89},{"nf":"122155","dt":"03/07/2026","vlr":1236.68}],"venc":[],"emit":null},{"g":"AENA","r":"BLOCO DE ONZE AEROPORTOS DO BRASIL S/A","e":"rkubota@aenabrasil.com.br","c":"Sao Paulo","uf":"SP","qtd":1,"aV":1402.97,"v30":0,"v60":0,"v90":0,"v91":0,"total":1402.97,"prox":[{"nf":"131383","dt":"10/07/2026","vlr":1402.97}],"venc":[],"emit":{"total":-161.06,"itens":[{"ref":"001135688-005","vlr":-161.06,"tipo":"ZV","obs":"Transf doc fatmto 90957191 COMUNICARE COMERCIO DE"}],"alerta":""}},{"g":"ALLTE 0364","r":"GUABI NUTRICAO E SAUDE ANIMAL SA","e":"Ludimila.Dutra@guabi.com.br","c":"Para de Minas","uf":"MG","qtd":28,"aV":114972.34,"v30":0,"v60":0,"v90":2243.82,"v91":0,"total":117216.16,"prox":[{"nf":"1270857","dt":"15/06/2026","vlr":2541.75},{"nf":"1272563","dt":"16/06/2026","vlr":6139.59},{"nf":"1273490","dt":"17/06/2026","vlr":7667.5}],"venc":[{"nf":"1225468","dt":"25/03/2026","vlr":2243.82,"d":80}],"emit":{"total":-2427.98,"itens":[{"ref":"001225468-005","vlr":-2243.82,"tipo":"RV","obs":"Transf doc fatmto 91110550 GUABI NUTRICAO E SAUDE"},{"ref":"03032026","vlr":-172.24,"tipo":"DA","obs":"Documento do Cliente 3032026 ERNST E YOUNG ASSESSO"},{"ref":"03112025","vlr":-11.92,"tipo":"DA","obs":"Documento do Cliente 3112025 ERNST E YOUNG ASSESSO"}],"alerta":"⚠️ NF 1225468 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"APERAM INOX","r":"APERAM INOX TUBOS BRASIL LTDA","e":"marcela.gomes@aperam.com","c":"Ribeirao Pires","uf":"SP","qtd":9,"aV":13629.62,"v30":0,"v60":0,"v90":0,"v91":0,"total":13629.62,"prox":[{"nf":"129192","dt":"07/07/2026","vlr":469.82},{"nf":"128575","dt":"07/07/2026","vlr":1071.0},{"nf":"129014","dt":"07/07/2026","vlr":1656.86}],"venc":[],"emit":{"total":-454.63,"itens":[{"ref":"30092025","vlr":-454.63,"tipo":"DA","obs":"Documento do Cliente 30092025 APERAM INOX AMERICA"}],"alerta":""}},{"g":"APPLUS 0142","r":"APPLUS QUALITEC SERVICOS DE ENGENHARIA LTDA","e":"amanda.costa@applus.com","c":"Varzea Grande","uf":"MT","qtd":14,"aV":43564.96,"v30":0,"v60":0,"v90":0,"v91":0,"total":43564.96,"prox":[{"nf":"1263368","dt":"19/06/2026","vlr":4338.24},{"nf":"1270830","dt":"25/06/2026","vlr":7792.91},{"nf":"1270767","dt":"25/06/2026","vlr":4513.7}],"venc":[],"emit":null},{"g":"APTIV B 0150","r":"APTIV MANUFATURA E SERVICOS DE DISTRIBUICAO LTDA","e":"atendimento.aptiv@brsupply.com.br","c":"Espirito Santo do Pinhal","uf":"SP","qtd":16,"aV":30212.79,"v30":0,"v60":0,"v90":0,"v91":0,"total":30212.79,"prox":[{"nf":"102626","dt":"15/06/2026","vlr":764.19},{"nf":"102860","dt":"15/06/2026","vlr":4627.85},{"nf":"1232644","dt":"30/06/2026","vlr":2500.95}],"venc":[],"emit":{"total":-2427.98,"itens":[{"ref":"001225468-005","vlr":-2243.82,"tipo":"RV","obs":"Transf doc fatmto 91110550 GUABI NUTRICAO E SAUDE"},{"ref":"03032026","vlr":-172.24,"tipo":"DA","obs":"Documento do Cliente 3032026 ERNST E YOUNG ASSESSO"},{"ref":"03112025","vlr":-11.92,"tipo":"DA","obs":"Documento do Cliente 3112025 ERNST E YOUNG ASSESSO"}],"alerta":""}},{"g":"BARUERI 0148","r":"VONNY COSMETICOS LTDA","e":"Maria.francielly@casadolojista.com.br","c":"Fortaleza","uf":"CE","qtd":35,"aV":25774.05,"v30":0,"v60":0,"v90":0,"v91":0,"total":25774.05,"prox":[{"nf":"1274895","dt":"18/06/2026","vlr":812.72},{"nf":"1279879","dt":"23/06/2026","vlr":982.0},{"nf":"129178","dt":"23/06/2026","vlr":340.11}],"venc":[],"emit":{"total":-738.54,"itens":[{"ref":"14042026","vlr":-284.1,"tipo":"DA","obs":"Documento do Cliente 14042026 MESSES GASES LTDA"},{"ref":"000061402-007","vlr":-454.44,"tipo":"RV","obs":"Transf doc fatmto 90927595 BAXTER HOSPITALAR LTDA"}],"alerta":""}},{"g":"BAXTER","r":"BAXTER HOSPITALAR LTDA","e":"isabella_brito@baxter.com","c":"Sao Paulo","uf":"SP","qtd":17,"aV":15923.41,"v30":0,"v60":0,"v90":0,"v91":454.44,"total":16377.85,"prox":[{"nf":"102736","dt":"15/06/2026","vlr":302.7},{"nf":"107390","dt":"23/06/2026","vlr":538.2},{"nf":"108055","dt":"25/06/2026","vlr":1168.3}],"venc":[{"nf":"61402","dt":"09/02/2026","vlr":454.44,"d":124}],"emit":{"total":-738.54,"itens":[{"ref":"14042026","vlr":-284.1,"tipo":"DA","obs":"Documento do Cliente 14042026 MESSES GASES LTDA"},{"ref":"000061402-007","vlr":-454.44,"tipo":"RV","obs":"Transf doc fatmto 90927595 BAXTER HOSPITALAR LTDA"}],"alerta":"⚠️ NF 61402 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"BAYER S 2169","r":"BAYER S/A","e":"atendimento.bayer@brsupply.com.br","c":"Sao Paulo","uf":"SP","qtd":47,"aV":43264.3,"v30":5900.64,"v60":0,"v90":0,"v91":0,"total":49164.94,"prox":[{"nf":"128091","dt":"15/06/2026","vlr":863.6},{"nf":"103726","dt":"16/06/2026","vlr":1512.72},{"nf":"129295","dt":"18/06/2026","vlr":1435.0}],"venc":[{"nf":"1213035","dt":"11/06/2026","vlr":2688.94,"d":2},{"nf":"126962","dt":"12/06/2026","vlr":522.76,"d":1}],"emit":null},{"g":"BRISTOL 0107","r":"BRISTOL-MYERS SQUIBB FARMACEUTICA LTDA","e":"servicosinternos@bms.com","c":"Sao Paulo","uf":"SP","qtd":3,"aV":2290.97,"v30":0,"v60":0,"v90":0,"v91":0,"total":2290.97,"prox":[{"nf":"120663","dt":"29/06/2026","vlr":623.21},{"nf":"124703","dt":"07/07/2026","vlr":553.79},{"nf":"128786","dt":"17/07/2026","vlr":1113.97}],"venc":[],"emit":null},{"g":"CAMARAD 0172","r":"CAMARADA ADMINISTRACAO DE RESTAURANTES SA","e":"tiago.guimaraes@camaradacamarao.com.br","c":"Mogi das Cruzes","uf":"SP","qtd":93,"aV":87330.82,"v30":1875.82,"v60":0,"v90":0,"v91":0,"total":89206.64,"prox":[{"nf":"115347","dt":"15/06/2026","vlr":480.22},{"nf":"1246494","dt":"15/06/2026","vlr":1145.87},{"nf":"1246480","dt":"15/06/2026","vlr":1145.87}],"venc":[{"nf":"1229356","dt":"29/05/2026","vlr":1098.18,"d":15},{"nf":"1257818","dt":"12/06/2026","vlr":777.64,"d":1}],"emit":{"total":-161.06,"itens":[{"ref":"001135688-005","vlr":-161.06,"tipo":"ZV","obs":"Transf doc fatmto 90957191 COMUNICARE COMERCIO DE"}],"alerta":""}},{"g":"CLARIANT","r":"CLARIANT BRASIL LTDA","e":"atendimento.clariant@brsupply.com.br","c":"Sao Paulo","uf":"SP","qtd":37,"aV":30101.51,"v30":929.0,"v60":0,"v90":0,"v91":0,"total":31030.51,"prox":[{"nf":"116038","dt":"15/06/2026","vlr":1067.41},{"nf":"116970","dt":"16/06/2026","vlr":1809.94},{"nf":"117571","dt":"19/06/2026","vlr":567.98}],"venc":[{"nf":"114555","dt":"12/06/2026","vlr":464.5,"d":1}],"emit":{"total":-738.54,"itens":[{"ref":"14042026","vlr":-284.1,"tipo":"DA","obs":"Documento do Cliente 14042026 MESSES GASES LTDA"},{"ref":"000061402-007","vlr":-454.44,"tipo":"RV","obs":"Transf doc fatmto 90927595 BAXTER HOSPITALAR LTDA"}],"alerta":""}},{"g":"COMUNIC 0134","r":"COMUNICARE COMERCIO DE APARELHOS AUDITIVOS LTDA","e":"yara.ferro@wsa.com","c":"Santos","uf":"SP","qtd":185,"aV":96814.99,"v30":0,"v60":0,"v90":0,"v91":0,"total":96814.99,"prox":[{"nf":"1247969","dt":"15/06/2026","vlr":1841.62},{"nf":"117066","dt":"16/06/2026","vlr":257.35},{"nf":"1250654","dt":"16/06/2026","vlr":247.7}],"venc":[],"emit":{"total":-161.06,"itens":[{"ref":"001135688-005","vlr":-161.06,"tipo":"ZV","obs":"Transf doc fatmto 90957191 COMUNICARE COMERCIO DE"}],"alerta":""}},{"g":"DANFOSS","r":"DANFOSS POWER SOLUTIONS INDUSTRIA E COMERCIO ELETROHIDRAULICA LTDA","e":"atendimento.danfoss@brsupply.com.br","c":"Caxias do Sul","uf":"RS","qtd":8,"aV":6861.55,"v30":0,"v60":0,"v90":0,"v91":0,"total":6861.55,"prox":[{"nf":"679673","dt":"24/06/2026","vlr":1167.96},{"nf":"691704","dt":"08/07/2026","vlr":1309.28},{"nf":"693867","dt":"15/07/2026","vlr":685.34}],"venc":[],"emit":null},{"g":"DENGO C 0371","r":"DENGO CHOCOLATES S.A.","e":"atendimento.dengo@brsupply.com.br","c":"Curitiba","uf":"PR","qtd":10,"aV":5410.83,"v30":0,"v60":0,"v90":0,"v91":0,"total":5410.83,"prox":[{"nf":"1281668","dt":"18/06/2026","vlr":562.26},{"nf":"130366","dt":"18/06/2026","vlr":466.29},{"nf":"130419","dt":"18/06/2026","vlr":545.24}],"venc":[],"emit":null},{"g":"ERNST YOUNG","r":"ERNST E YOUNG ASSESSORIA EMPRESARIAL LTDA","e":"paula.souza.ext@br.ey.com","c":"Sao Paulo","uf":"SP","qtd":22,"aV":26581.82,"v30":9400.91,"v60":0,"v90":0,"v91":0,"total":35982.73,"prox":[{"nf":"121511","dt":"15/06/2026","vlr":7227.15},{"nf":"125299","dt":"25/06/2026","vlr":3481.42},{"nf":"126178","dt":"26/06/2026","vlr":507.39}],"venc":[{"nf":"112664","dt":"25/05/2026","vlr":4024.02,"d":19},{"nf":"1238628","dt":"25/05/2026","vlr":610.16,"d":19},{"nf":"1240738","dt":"25/05/2026","vlr":453.44,"d":19},{"nf":"1238508","dt":"25/05/2026","vlr":1903.97,"d":19}],"emit":{"total":-2427.98,"itens":[{"ref":"001225468-005","vlr":-2243.82,"tipo":"RV","obs":"Transf doc fatmto 91110550 GUABI NUTRICAO E SAUDE"},{"ref":"03032026","vlr":-172.24,"tipo":"DA","obs":"Documento do Cliente 3032026 ERNST E YOUNG ASSESSO"},{"ref":"03112025","vlr":-11.92,"tipo":"DA","obs":"Documento do Cliente 3112025 ERNST E YOUNG ASSESSO"}],"alerta":""}},{"g":"FIAT","r":"STELLANTIS AUTOMOVEIS BRASIL LTDA","e":"atendimento.fiat@brsupply.com.br","c":"Betim","uf":"MG","qtd":5,"aV":0,"v30":0,"v60":0,"v90":0,"v91":4108.4,"total":4108.4,"prox":[],"venc":[{"nf":"713130","dt":"16/06/2025","vlr":678.85,"d":362},{"nf":"908180","dt":"31/07/2025","vlr":127.71,"d":317},{"nf":"885441","dt":"20/08/2025","vlr":123.04,"d":297},{"nf":"946565","dt":"01/09/2025","vlr":401.4,"d":285}],"emit":null},{"g":"FLEXTRO 0110","r":"FLEXTRONICS DA AMAZONIA LTDA","e":"","c":"Manaus","uf":"AM","qtd":1,"aV":5100.0,"v30":0,"v60":0,"v90":0,"v91":0,"total":5100.0,"prox":[{"nf":"1272045","dt":"06/07/2026","vlr":5100.0}],"venc":[],"emit":null},{"g":"MODERN","r":"MODERN TRANSPORTE AEREO DE CARGA S/A","e":"","c":"Recife","uf":"PE","qtd":1,"aV":0,"v30":1020.32,"v60":0,"v90":0,"v91":0,"total":1020.32,"prox":[],"venc":[{"nf":"1264399","dt":"03/06/2026","vlr":1020.32,"d":10}],"emit":null},{"g":"MODERN 1510","r":"MODERN TRANSPORTE AEREO DE CARGA S.A.","e":"","c":"Goiania","uf":"GO","qtd":1,"aV":483.38,"v30":0,"v60":0,"v90":0,"v91":0,"total":483.38,"prox":[{"nf":"1285586","dt":"23/06/2026","vlr":483.38}],"venc":[],"emit":null},{"g":"MODINE 0241","r":"MODINE DO BRASIL SISTEMAS TERMICOS LTDA.","e":"leandro.c.santos@modine.com","c":"Guarulhos","uf":"SP","qtd":5,"aV":4724.98,"v30":0,"v60":0,"v90":0,"v91":0,"total":4724.98,"prox":[{"nf":"125451","dt":"15/06/2026","vlr":547.69},{"nf":"127495","dt":"18/06/2026","vlr":509.98},{"nf":"129461","dt":"23/06/2026","vlr":567.11}],"venc":[],"emit":null},{"g":"MRS SP","r":"MRS LOGISTICA S/A","e":"","c":"Sao Paulo","uf":"SP","qtd":3,"aV":6581.05,"v30":3232.8,"v60":0,"v90":0,"v91":0,"total":9813.85,"prox":[{"nf":"1293349","dt":"03/08/2026","vlr":6581.05}],"venc":[{"nf":"106347","dt":"22/05/2026","vlr":1027.45,"d":22},{"nf":"1227521","dt":"26/05/2026","vlr":2205.35,"d":18}],"emit":null},{"g":"PQ PINHEIROS","r":"PARQUE DOS PINHEIROS ADMINISTRACAO DE BENS LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":4,"aV":1427.2,"v30":0,"v60":0,"v90":0,"v91":0,"total":1427.2,"prox":[{"nf":"133217","dt":"29/06/2026","vlr":356.8}],"venc":[],"emit":null},{"g":"PREVISUL","r":"ODONTO EMPRESAS CONVENIOS DENTARIOS LTDA","e":"Jose.Oliveira.Stefanini@cnpseguradora.com.br","c":"Barueri","uf":"SP","qtd":3,"aV":12215.64,"v30":0,"v60":0,"v90":0,"v91":0,"total":12215.64,"prox":[{"nf":"135114","dt":"20/07/2026","vlr":5922.88},{"nf":"135436","dt":"20/07/2026","vlr":3146.38}],"venc":[],"emit":null},{"g":"SAIPEM 0604","r":"SAIPEM DO BRASIL SERVICOS DE PETROLEOLTDA","e":"juliana.cunha@saipem.com","c":"Guaruja","uf":"SP","qtd":4,"aV":53491.76,"v30":0,"v60":0,"v90":0,"v91":0,"total":53491.76,"prox":[{"nf":"124865","dt":"06/07/2026","vlr":11998.87},{"nf":"1287450","dt":"13/07/2026","vlr":14754.8},{"nf":"133179","dt":"13/07/2026","vlr":11983.29}],"venc":[],"emit":null},{"g":"SIMBIOS 0169","r":"SIMBIOSE INDUSTRIA E COMERCIO DE FERTILIZANTES E INSUMOS MICROBIOLOGICOS LTDA","e":"jose.nascimento@cognycsc.com.br","c":"Ribeirao Preto","uf":"SP","qtd":3,"aV":2530.46,"v30":715.08,"v60":0,"v90":0,"v91":0,"total":3245.54,"prox":[{"nf":"121868","dt":"16/06/2026","vlr":738.73},{"nf":"1264716","dt":"23/06/2026","vlr":1791.73}],"venc":[{"nf":"1246495","dt":"02/06/2026","vlr":715.08,"d":11}],"emit":{"total":-2427.98,"itens":[{"ref":"001225468-005","vlr":-2243.82,"tipo":"RV","obs":"Transf doc fatmto 91110550 GUABI NUTRICAO E SAUDE"},{"ref":"03032026","vlr":-172.24,"tipo":"DA","obs":"Documento do Cliente 3032026 ERNST E YOUNG ASSESSO"},{"ref":"03112025","vlr":-11.92,"tipo":"DA","obs":"Documento do Cliente 3112025 ERNST E YOUNG ASSESSO"}],"alerta":""}},{"g":"SWISSPORT","r":"CARGO SERVICE CENTER BRAZIL SERVICOS AUX","e":"atendimento.swisspor@brsupply.com.br","c":"Guarulhos","uf":"SP","qtd":37,"aV":94294.72,"v30":23251.41,"v60":0,"v90":0,"v91":3450.0,"total":120996.13,"prox":[{"nf":"1210223","dt":"30/06/2026","vlr":1692.35},{"nf":"124340","dt":"06/07/2026","vlr":650.0},{"nf":"124353","dt":"06/07/2026","vlr":1526.6}],"venc":[{"nf":"1146690","dt":"13/03/2026","vlr":3450.0,"d":92},{"nf":"114360","dt":"12/06/2026","vlr":3315.92,"d":1},{"nf":"1242887","dt":"12/06/2026","vlr":5265.24,"d":1},{"nf":"1244011","dt":"12/06/2026","vlr":2934.05,"d":1}],"emit":{"total":-19.8,"itens":[{"ref":"21102025","vlr":-19.8,"tipo":"DA","obs":"Documento do Cliente 21102025 CARGO SERVICE CENTER"}],"alerta":""}},{"g":"THOMSON","r":"THOMSON REUTERS BRASIL CONTEUDO E TECNOLOGIA LTDA","e":"lucielenamaria.sorana@thomsonreuters.com","c":"Campinas","uf":"SP","qtd":5,"aV":6956.17,"v30":0,"v60":0,"v90":0,"v91":0,"total":6956.17,"prox":[{"nf":"137565","dt":"09/07/2026","vlr":1498.14},{"nf":"137360","dt":"09/07/2026","vlr":598.83},{"nf":"138318","dt":"10/07/2026","vlr":815.72}],"venc":[],"emit":null},{"g":"VALID 0147","r":"VALID SOLUCOES S A","e":"ariane.jadanhi@valid.com","c":"Sorocaba","uf":"SP","qtd":63,"aV":95463.18,"v30":1208.23,"v60":0,"v90":0,"v91":0,"total":96671.41,"prox":[{"nf":"112646","dt":"15/06/2026","vlr":966.28},{"nf":"1242677","dt":"15/06/2026","vlr":4924.89},{"nf":"1242039","dt":"15/06/2026","vlr":3362.64}],"venc":[{"nf":"1236918","dt":"08/06/2026","vlr":1208.23,"d":5}],"emit":null},{"g":"XP INVEST","r":"XP EVENTOS LTDA","e":"elivania.silva.t8208@xpi.com.br","c":"Sao Paulo","uf":"SP","qtd":29,"aV":25227.24,"v30":2888.62,"v60":0,"v90":0,"v91":0,"total":28115.86,"prox":[{"nf":"112755","dt":"15/06/2026","vlr":981.6},{"nf":"1270374","dt":"15/07/2026","vlr":1034.05},{"nf":"1270326","dt":"15/07/2026","vlr":1776.19}],"venc":[{"nf":"1202949","dt":"15/05/2026","vlr":735.8,"d":29},{"nf":"1204547","dt":"15/05/2026","vlr":1257.97,"d":29},{"nf":"97316","dt":"15/05/2026","vlr":894.85,"d":29}],"emit":{"total":-3097.22,"itens":[{"ref":"05062026","vlr":-402.1,"tipo":"DA","obs":"Documento do Cliente 5062026 XP INVESTIMENTOS CORR"},{"ref":"06042026","vlr":-1256.06,"tipo":"DA","obs":"CH 1758224"},{"ref":"27032026","vlr":-24.67,"tipo":"DA","obs":"Documento do Cliente 27032026 XP INVESTIMENTOS COR"}],"alerta":""}},{"g":"YPF BRA 0105","r":"USIBLEND INDUSTRIA E SERVICOS LTDA","e":"paloma.almeida@usiblend.com.br","c":"Diadema","uf":"SP","qtd":1,"aV":7066.65,"v30":0,"v60":0,"v90":0,"v91":0,"total":7066.65,"prox":[{"nf":"138584","dt":"10/07/2026","vlr":7066.65}],"venc":[],"emit":{"total":-2427.98,"itens":[{"ref":"001225468-005","vlr":-2243.82,"tipo":"RV","obs":"Transf doc fatmto 91110550 GUABI NUTRICAO E SAUDE"},{"ref":"03032026","vlr":-172.24,"tipo":"DA","obs":"Documento do Cliente 3032026 ERNST E YOUNG ASSESSO"},{"ref":"03112025","vlr":-11.92,"tipo":"DA","obs":"Documento do Cliente 3112025 ERNST E YOUNG ASSESSO"}],"alerta":""}}],"Jessica Antunes da Silva":[{"g":"AASP","r":"ASSOCIACAO DOS ADVOGADOS DE SAO PAULO","e":"compras@aasp.org.br","c":"Sao Paulo","uf":"SP","qtd":3,"aV":16246.01,"v30":0,"v60":0,"v90":0,"v91":0,"total":16246.01,"prox":[{"nf":"123262","dt":"30/06/2026","vlr":1334.97},{"nf":"131381","dt":"30/06/2026","vlr":7455.52}],"venc":[],"emit":null},{"g":"ABBOTT","r":"ABBOTT LABORATORIOS DO BRASIL LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":6,"aV":3755.69,"v30":0,"v60":0,"v90":0,"v91":0,"total":3755.69,"prox":[{"nf":"111657","dt":"06/07/2026","vlr":249.0},{"nf":"117180","dt":"16/07/2026","vlr":633.5},{"nf":"121663","dt":"29/07/2026","vlr":261.3}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"ABBOTT 1279","r":"ABBOTT LABORATORIOS DO BRASIL LTDA","e":"","c":"Rio de Janeiro","uf":"RJ","qtd":2,"aV":2048.8,"v30":0,"v60":0,"v90":0,"v91":0,"total":2048.8,"prox":[{"nf":"1278983","dt":"17/08/2026","vlr":1024.4}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"ABBOTT 3212","r":"ABBOTT LABORATORIOS DO BRASIL LTDA","e":"","c":"Itapevi","uf":"SP","qtd":2,"aV":4019.6,"v30":0,"v60":0,"v90":0,"v91":0,"total":4019.6,"prox":[{"nf":"111436","dt":"06/07/2026","vlr":2009.8}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"ABBVIE","r":"ABBVIE FARMACEUTICA LTDA.","e":"rafael.sobral@abbvie.com","c":"Sao Paulo","uf":"SP","qtd":1,"aV":1331.58,"v30":0,"v60":0,"v90":0,"v91":0,"total":1331.58,"prox":[{"nf":"126795","dt":"13/07/2026","vlr":1331.58}],"venc":[],"emit":null},{"g":"ABBVIE 0230","r":"ABBVIE FARMACEUTICA LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":2,"aV":1492.75,"v30":0,"v60":0,"v90":0,"v91":124.33,"total":1617.08,"prox":[{"nf":"111915","dt":"06/07/2026","vlr":1492.75}],"venc":[{"nf":"964147","dt":"28/10/2025","vlr":124.33,"d":228}],"emit":{"total":-26.31,"itens":[{"ref":"001078509-005","vlr":-26.31,"tipo":"ZV","obs":"Transf doc fatmto 90862800 CONSTRUTORA ELEVACAO LT"}],"alerta":""}},{"g":"AFAM 0001","r":"ASSOCIACAO AFAM DE ASSISTENCIA FARMACEUTICA","e":"taubate@afam.com.br","c":"Taubate","uf":"SP","qtd":12,"aV":12102.33,"v30":0,"v60":0,"v90":0,"v91":0,"total":12102.33,"prox":[{"nf":"131374","dt":"15/06/2026","vlr":1136.81},{"nf":"131379","dt":"15/06/2026","vlr":478.28},{"nf":"131315","dt":"15/06/2026","vlr":792.85}],"venc":[],"emit":{"total":-45.8,"itens":[{"ref":"12112025","vlr":-45.8,"tipo":"ZV","obs":"Chamado 1649904"}],"alerta":""}},{"g":"AGASUS","r":"VOKE S.A.","e":"barbara.santos@agasus.com.br","c":"Sao Paulo","uf":"SP","qtd":4,"aV":1067.4,"v30":0,"v60":0,"v90":0,"v91":5749.54,"total":6816.94,"prox":[{"nf":"1292132","dt":"01/07/2026","vlr":1067.4}],"venc":[{"nf":"786404","dt":"14/04/2025","vlr":813.4,"d":425},{"nf":"813663","dt":"30/04/2025","vlr":585.9,"d":409},{"nf":"855716","dt":"06/06/2025","vlr":4350.24,"d":372}],"emit":null},{"g":"AUTOAMERICAN","r":"AUTO AMERICANO SA DISTRIBUIDOR DE PECAS","e":"teadm2@autoamericano.com.br","c":"Teresina","uf":"PI","qtd":4,"aV":3342.07,"v30":708.85,"v60":0,"v90":0,"v91":0,"total":4050.92,"prox":[{"nf":"1284836","dt":"23/06/2026","vlr":816.88},{"nf":"133332","dt":"25/06/2026","vlr":1527.36},{"nf":"1305134","dt":"10/07/2026","vlr":997.83}],"venc":[{"nf":"1267647","dt":"12/06/2026","vlr":708.85,"d":1}],"emit":null},{"g":"AXALTA 0145","r":"AXALTA COATING SYSTEMS BRASIL LTDA.","e":"","c":"Guarulhos","uf":"SP","qtd":8,"aV":20139.77,"v30":0,"v60":0,"v90":0,"v91":0,"total":20139.77,"prox":[{"nf":"121650","dt":"15/06/2026","vlr":11821.85},{"nf":"121669","dt":"15/06/2026","vlr":545.24},{"nf":"132455","dt":"13/07/2026","vlr":1193.28}],"venc":[],"emit":null},{"g":"BAMCAF 0161","r":"VB TRANSPORTES E TURISMO LTDA.","e":"estoque.vb3@grupobelarmino.com.br","c":"Campinas","uf":"SP","qtd":13,"aV":13932.06,"v30":2988.22,"v60":2180.81,"v90":0,"v91":0,"total":19101.09,"prox":[{"nf":"125075","dt":"15/06/2026","vlr":876.48},{"nf":"125406","dt":"15/06/2026","vlr":976.97},{"nf":"126691","dt":"17/06/2026","vlr":2424.3}],"venc":[{"nf":"110746","dt":"11/05/2026","vlr":2180.81,"d":33},{"nf":"122308","dt":"08/06/2026","vlr":1465.25,"d":5},{"nf":"124851","dt":"12/06/2026","vlr":1522.97,"d":1}],"emit":{"total":-1866.97,"itens":[{"ref":"000108364-007","vlr":-1866.97,"tipo":"RV","obs":"Transf doc fatmto 91115171 HEATING E COOLING TECNO"}],"alerta":""}},{"g":"BLUE AN 0180","r":"BLUE ANGELS SEGURANCA PRIVADA E TRANSPORTE DE VALORES LTDA","e":"edmop@blueangels.com.br","c":"Sao Paulo","uf":"SP","qtd":7,"aV":0,"v30":0,"v60":0,"v90":0,"v91":11555.63,"total":11555.63,"prox":[],"venc":[{"nf":"703519","dt":"14/02/2025","vlr":1343.66,"d":484},{"nf":"714175","dt":"17/02/2025","vlr":1996.54,"d":481},{"nf":"730269","dt":"28/02/2025","vlr":1477.98,"d":470},{"nf":"745264","dt":"13/03/2025","vlr":988.87,"d":457}],"emit":null},{"g":"CENTRAL0109","r":"SPE SOROCABA AMBIENTAL S.A.","e":"bfsobral@orizonvr.com.br","c":"Sorocaba","uf":"SP","qtd":10,"aV":22981.26,"v30":0,"v60":0,"v90":0,"v91":0,"total":22981.26,"prox":[{"nf":"1261032","dt":"29/06/2026","vlr":2612.23},{"nf":"121855","dt":"30/06/2026","vlr":976.85},{"nf":"1265950","dt":"06/07/2026","vlr":1044.89}],"venc":[],"emit":null},{"g":"CON ELEVACAO","r":"CONSTRUTORA ELEVACAO LTDA","e":"gislaine.lima@construtoraelevacao.com.br","c":"Contagem","uf":"MG","qtd":1,"aV":0,"v30":0,"v60":0,"v90":0,"v91":26.31,"total":26.31,"prox":[],"venc":[{"nf":"1078509","dt":"05/11/2025","vlr":26.31,"d":220}],"emit":{"total":-26.31,"itens":[{"ref":"001078509-005","vlr":-26.31,"tipo":"ZV","obs":"Transf doc fatmto 90862800 CONSTRUTORA ELEVACAO LT"}],"alerta":"⚠️ NF 1078509 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"CORSUL 0409","r":"CORSUL COMERCIO E REPRESENTACOES DO SUL LTDA","e":"","c":"Caxias do Sul","uf":"RS","qtd":2,"aV":3764.05,"v30":0,"v60":0,"v90":0,"v91":0,"total":3764.05,"prox":[{"nf":"737478","dt":"06/07/2026","vlr":1608.19},{"nf":"737841","dt":"08/07/2026","vlr":2155.86}],"venc":[],"emit":{"total":-1912.77,"itens":[{"ref":"000108364-007","vlr":-1866.97,"tipo":"RV","obs":"Transf doc fatmto 91115171 HEATING E COOLING TECNO"},{"ref":"12112025","vlr":-45.8,"tipo":"ZV","obs":"Chamado 1649904"}],"alerta":""}},{"g":"CPACOMPRODAG","r":"CMC FERRARIN COMERCIO DE MAQUINAS E IMPLEMENTOS AGRICOLAS LTDA","e":"coord.adm-santamaria@cmimaquinas.com.br","c":"Santa Maria","uf":"RS","qtd":1,"aV":0,"v30":0,"v60":0,"v90":0,"v91":602.64,"total":602.64,"prox":[],"venc":[{"nf":"590913","dt":"05/12/2025","vlr":602.64,"d":190}],"emit":{"total":-45.8,"itens":[{"ref":"12112025","vlr":-45.8,"tipo":"ZV","obs":"Chamado 1649904"}],"alerta":""}},{"g":"ELECNOR","r":"ELECNOR DO BRASIL LTDA","e":"eric.bernardo@elecnor.es","c":"Estiva Gerbi","uf":"SP","qtd":14,"aV":78323.5,"v30":0,"v60":0,"v90":0,"v91":0,"total":78323.5,"prox":[{"nf":"121763","dt":"29/06/2026","vlr":5187.5},{"nf":"121770","dt":"29/06/2026","vlr":6945.56},{"nf":"121551","dt":"29/06/2026","vlr":23538.69}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"EMERSON 0100","r":"EMERSON PROCESS MANAGEMENT LTDA","e":"giseli.dona@emerson.com","c":"Sorocaba","uf":"SP","qtd":11,"aV":25370.84,"v30":0,"v60":0,"v90":0,"v91":0,"total":25370.84,"prox":[{"nf":"115859","dt":"17/06/2026","vlr":1293.87},{"nf":"117009","dt":"17/06/2026","vlr":5038.8},{"nf":"121774","dt":"01/07/2026","vlr":2101.44}],"venc":[],"emit":null},{"g":"Farmaci 0161","r":"FARMACIA DO PEDRINHO LTDA","e":"filial7cd@pedrinhofarmacias.com.br","c":"Santa Rosa","uf":"RS","qtd":11,"aV":2503.73,"v30":0,"v60":0,"v90":0,"v91":0,"total":2503.73,"prox":[{"nf":"739234","dt":"09/07/2026","vlr":264.13},{"nf":"739387","dt":"09/07/2026","vlr":222.46},{"nf":"739443","dt":"09/07/2026","vlr":205.12}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"GREAT W 0155","r":"GREAT WALL MOTOR BRASIL LTDA","e":"","c":"Iracemapolis","uf":"SP","qtd":8,"aV":8567.24,"v30":0,"v60":0,"v90":0,"v91":0,"total":8567.24,"prox":[{"nf":"135427","dt":"03/07/2026","vlr":840.0},{"nf":"135824","dt":"03/07/2026","vlr":407.95},{"nf":"135829","dt":"03/07/2026","vlr":1444.16}],"venc":[],"emit":null},{"g":"GREAT WALL M","r":"GREAT WALL MOTOR BRASIL LTDA","e":"suporte.greatwall@brsupply.com.br","c":"Sao Paulo","uf":"SP","qtd":7,"aV":12355.45,"v30":1683.53,"v60":0,"v90":6734.12,"v91":0,"total":20773.1,"prox":[{"nf":"135560","dt":"03/07/2026","vlr":5690.0},{"nf":"135850","dt":"03/07/2026","vlr":6665.45}],"venc":[{"nf":"97525","dt":"03/04/2026","vlr":1683.53,"d":71},{"nf":"103488","dt":"11/06/2026","vlr":1683.53,"d":2}],"emit":null},{"g":"HEATIN 0120","r":"HEATING E COOLING TECNOLOGIA TERMICA LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":1,"aV":0,"v30":0,"v60":0,"v90":1866.97,"v91":0,"total":1866.97,"prox":[],"venc":[{"nf":"108364","dt":"27/03/2026","vlr":1866.97,"d":78}],"emit":{"total":-1866.97,"itens":[{"ref":"000108364-007","vlr":-1866.97,"tipo":"RV","obs":"Transf doc fatmto 91115171 HEATING E COOLING TECNO"}],"alerta":"⚠️ NF 108364 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"IMPR 0178","r":"IMPRELL - EDITORA GRAFICA LTDA","e":"alan@imprell.com.br","c":"Arvorezinha","uf":"RS","qtd":1,"aV":696.08,"v30":0,"v60":0,"v90":0,"v91":0,"total":696.08,"prox":[{"nf":"1294902","dt":"01/07/2026","vlr":696.08}],"venc":[],"emit":null},{"g":"Jingdon 0113","r":"JINGDONG INDUSTRIALS DO BRASIL LTDA","e":"atendimento.jd@brsupply.com.br","c":"Sao Paulo","uf":"SP","qtd":203,"aV":78263.98,"v30":25948.1,"v60":34693.95,"v90":140509.79,"v91":425.35,"total":279841.17,"prox":[{"nf":"116010","dt":"15/06/2026","vlr":137.79},{"nf":"115755","dt":"15/06/2026","vlr":86.72},{"nf":"115841","dt":"15/06/2026","vlr":290.46}],"venc":[{"nf":"31655","dt":"04/11/2025","vlr":425.35,"d":221},{"nf":"1133605","dt":"06/04/2026","vlr":232.32,"d":68},{"nf":"1129849","dt":"06/04/2026","vlr":288.43,"d":68},{"nf":"1125616","dt":"06/04/2026","vlr":2624.38,"d":68}],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":"⚠️ NF 31655 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"KPMG 0129","r":"KPMG AUDITORES INDEPENDENTES LTDA","e":"rsergio@kpmg.com.br","c":"Sao Paulo","uf":"SP","qtd":21,"aV":48875.82,"v30":0,"v60":0,"v90":0,"v91":0,"total":48875.82,"prox":[{"nf":"121782","dt":"29/06/2026","vlr":793.5},{"nf":"122461","dt":"03/07/2026","vlr":8221.33},{"nf":"123184","dt":"06/07/2026","vlr":1280.68}],"venc":[],"emit":null},{"g":"LOCADOR 0190","r":"LOCADORA DL DO BRASIL LTDA.","e":"camila.testai@deutsche-leasing.com.br","c":"Sao Paulo","uf":"SP","qtd":1,"aV":1797.4,"v30":0,"v60":0,"v90":0,"v91":0,"total":1797.4,"prox":[{"nf":"133980","dt":"29/06/2026","vlr":1797.4}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"MASTER 0129","r":"MASTER VIGILANCIA ESPECIALIZADA LTDA","e":"ana.b@grupoaltum.com.br","c":"Londrina","uf":"PR","qtd":14,"aV":12387.32,"v30":0,"v60":0,"v90":0,"v91":0,"total":12387.32,"prox":[{"nf":"1281320","dt":"17/06/2026","vlr":366.56},{"nf":"724255","dt":"18/06/2026","vlr":1017.7},{"nf":"724445","dt":"18/06/2026","vlr":1973.46}],"venc":[],"emit":null},{"g":"MOTORMAC","r":"DISTRIBUIDORA MERIDIONAL DE MOTORES CUMMINS LTDA","e":"daniel.jesus@motormac.com.br","c":"Chapeco","uf":"SC","qtd":40,"aV":38099.32,"v30":1062.51,"v60":0,"v90":1000.79,"v91":590.88,"total":40753.5,"prox":[{"nf":"702082","dt":"23/06/2026","vlr":454.96},{"nf":"701938","dt":"23/06/2026","vlr":1306.13},{"nf":"701009","dt":"23/06/2026","vlr":519.85}],"venc":[{"nf":"593149","dt":"23/01/2026","vlr":590.88,"d":141},{"nf":"636578","dt":"23/03/2026","vlr":1000.79,"d":82},{"nf":"664947","dt":"25/05/2026","vlr":1062.51,"d":19}],"emit":{"total":-45.8,"itens":[{"ref":"12112025","vlr":-45.8,"tipo":"ZV","obs":"Chamado 1649904"}],"alerta":""}},{"g":"PETRONA 0186","r":"PETRONAS LUBRIFICANTES BRASIL SA","e":"","c":"Contagem","uf":"MG","qtd":4,"aV":13605.35,"v30":0,"v60":0,"v90":0,"v91":0,"total":13605.35,"prox":[{"nf":"1280187","dt":"28/07/2026","vlr":3648.69},{"nf":"1303486","dt":"28/07/2026","vlr":3612.14},{"nf":"1305372","dt":"28/07/2026","vlr":2695.83}],"venc":[],"emit":{"total":-130.13,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"}],"alerta":""}},{"g":"PWC","r":"PRICEWATERHOUSECOOPERS AUDITORES INDEPENDENTES LTDA","e":"ivana.x.ribeiro@pwc.com","c":"Fortaleza","uf":"CE","qtd":2,"aV":1192.82,"v30":0,"v60":0,"v90":0,"v91":0,"total":1192.82,"prox":[{"nf":"1284718","dt":"22/06/2026","vlr":739.83},{"nf":"133154","dt":"25/06/2026","vlr":452.99}],"venc":[],"emit":null},{"g":"RECOFAR 0602","r":"RECOFARMA INDUSTRIA DO AMAZONAS LTDA","e":"macardoso@coca-cola.com","c":"Rio de Janeiro","uf":"RJ","qtd":10,"aV":15275.49,"v30":0,"v60":0,"v90":0,"v91":0,"total":15275.49,"prox":[{"nf":"1259245","dt":"29/06/2026","vlr":1060.5},{"nf":"1258270","dt":"29/06/2026","vlr":883.96},{"nf":"1289477","dt":"28/07/2026","vlr":3652.7}],"venc":[],"emit":{"total":-555.48,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"},{"ref":"000031655-007","vlr":-425.35,"tipo":"RV","obs":"Transf doc fatmto 90814820 JINGDONG INDUSTRIALS DO"}],"alerta":""}},{"g":"RIOGRANDEEMB","r":"RIO GRANDE EMBALAGENS LTDA.","e":"comercial1@riograndeembalagens.com","c":"Sao Leopoldo","uf":"RS","qtd":2,"aV":627.6,"v30":0,"v60":0,"v90":0,"v91":0,"total":627.6,"prox":[{"nf":"713028","dt":"06/07/2026","vlr":300.1},{"nf":"1298740","dt":"07/08/2026","vlr":327.5}],"venc":[],"emit":null},{"g":"SALLVE 0195","r":"SALLVE COMERCIO DE COSMETICOS LTDA","e":"contem1g-centernorte@sallve.com","c":"Sao Paulo","uf":"SP","qtd":13,"aV":12147.31,"v30":1580.38,"v60":1650.68,"v90":0,"v91":0,"total":15378.37,"prox":[{"nf":"114898","dt":"15/06/2026","vlr":1418.45},{"nf":"118329","dt":"22/06/2026","vlr":203.3},{"nf":"121448","dt":"29/06/2026","vlr":3508.52}],"venc":[{"nf":"1195161","dt":"27/04/2026","vlr":489.09,"d":47},{"nf":"1192459","dt":"27/04/2026","vlr":336.25,"d":47},{"nf":"1239249","dt":"08/06/2026","vlr":461.41,"d":5},{"nf":"1237232","dt":"08/06/2026","vlr":328.78,"d":5}],"emit":{"total":-45.8,"itens":[{"ref":"12112025","vlr":-45.8,"tipo":"ZV","obs":"Chamado 1649904"}],"alerta":""}},{"g":"SONDA","r":"SONDA PROCWORK INFORMATICA LTDA","e":"eva.medeiros@sonda.com","c":"Santana de Parnaiba","uf":"SP","qtd":2,"aV":5393.3,"v30":0,"v60":0,"v90":0,"v91":0,"total":5393.3,"prox":[{"nf":"128530","dt":"17/06/2026","vlr":4619.37},{"nf":"1301532","dt":"10/07/2026","vlr":773.93}],"venc":[],"emit":{"total":-829.52,"itens":[{"ref":"05122025","vlr":-820.53,"tipo":"DA","obs":"Documento do Cliente 5122025 TECNOSET INFORMATICA"},{"ref":"05092025","vlr":-8.99,"tipo":"DA","obs":"Documento do Cliente 5092025 TECNOSET INFORMATICA"}],"alerta":""}},{"g":"SPAD CO 0162","r":"SPAD COMERCIO DE COSMETICOS LTDA","e":"adcos.eldorado@adcos.com.br","c":"Sao Paulo","uf":"SP","qtd":7,"aV":0,"v30":915.22,"v60":0,"v90":1035.82,"v91":1491.27,"total":3442.31,"prox":[],"venc":[{"nf":"28671","dt":"17/11/2025","vlr":586.15,"d":208},{"nf":"57643","dt":"15/01/2026","vlr":450.98,"d":149},{"nf":"57642","dt":"15/01/2026","vlr":454.14,"d":149},{"nf":"93555","dt":"09/04/2026","vlr":517.91,"d":65}],"emit":{"total":-45.8,"itens":[{"ref":"12112025","vlr":-45.8,"tipo":"ZV","obs":"Chamado 1649904"}],"alerta":""}},{"g":"SPENCER 0371","r":"SPENCER TRANSPORTES LTDA","e":"almoxarifado@d1transportes.com.br","c":"Sao Paulo","uf":"SP","qtd":9,"aV":3102.77,"v30":0,"v60":0,"v90":0,"v91":0,"total":3102.77,"prox":[{"nf":"105426","dt":"19/06/2026","vlr":342.88},{"nf":"131108","dt":"22/06/2026","vlr":269.01},{"nf":"139981","dt":"13/07/2026","vlr":572.92}],"venc":[],"emit":{"total":-26.31,"itens":[{"ref":"001078509-005","vlr":-26.31,"tipo":"ZV","obs":"Transf doc fatmto 90862800 CONSTRUTORA ELEVACAO LT"}],"alerta":""}},{"g":"Spencer 0533","r":"SPENCER TRANSPORTES LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":13,"aV":7117.75,"v30":0,"v60":0,"v90":0,"v91":0,"total":7117.75,"prox":[{"nf":"128804","dt":"18/06/2026","vlr":590.39},{"nf":"119747","dt":"23/06/2026","vlr":503.49},{"nf":"137226","dt":"08/07/2026","vlr":516.24}],"venc":[],"emit":{"total":-26.31,"itens":[{"ref":"001078509-005","vlr":-26.31,"tipo":"ZV","obs":"Transf doc fatmto 90862800 CONSTRUTORA ELEVACAO LT"}],"alerta":""}},{"g":"TECNOSET","r":"SCHAEFFLER BRASIL LTDA.","e":"atendimento.tecnoset@brsupply.com.br","c":"Embu das Artes","uf":"SP","qtd":10,"aV":14627.54,"v30":1659.9,"v60":0,"v90":0,"v91":0,"total":16287.44,"prox":[{"nf":"124055","dt":"15/07/2026","vlr":877.83},{"nf":"132084","dt":"27/07/2026","vlr":877.83},{"nf":"134973","dt":"05/08/2026","vlr":1574.84}],"venc":[{"nf":"111123","dt":"05/06/2026","vlr":1659.9,"d":8}],"emit":{"total":-130.13,"itens":[{"ref":"05052026","vlr":-130.13,"tipo":"DA","obs":"Documento do Cliente 5052026 ELECNOR DO BRASIL LTD"}],"alerta":""}},{"g":"THYSSEN 0102","r":"THYSSENKRUPP SPRINGS  STABILIZERS BRASIL LTDA.","e":"adriano.rezende@thyssenkrupp-springs-stabilizers.com","c":"Sao Paulo","uf":"SP","qtd":11,"aV":23728.08,"v30":0,"v60":0,"v90":0,"v91":0,"total":23728.08,"prox":[{"nf":"126844","dt":"25/06/2026","vlr":11108.36},{"nf":"1273825","dt":"25/06/2026","vlr":4092.47},{"nf":"1300718","dt":"27/07/2026","vlr":266.06}],"venc":[],"emit":null},{"g":"THYSSENKRUPP","r":"THYSSENKRUPP SPRINGS  STABILIZERS BRASIL LTDA","e":"","c":"Ibirite","uf":"MG","qtd":3,"aV":2423.4,"v30":0,"v60":0,"v90":0,"v91":0,"total":2423.4,"prox":[{"nf":"1293799","dt":"30/06/2026","vlr":807.8},{"nf":"1305291","dt":"30/06/2026","vlr":807.8},{"nf":"1302016","dt":"27/07/2026","vlr":807.8}],"venc":[],"emit":null},{"g":"TOTVS SP","r":"LEXOS SOLUCAO EM TECNOLOGIA LTDA","e":"estela.silva@totvs.com.br","c":"Sao Jose dos Campos","uf":"SP","qtd":20,"aV":31393.31,"v30":0,"v60":0,"v90":0,"v91":0,"total":31393.31,"prox":[{"nf":"122350","dt":"15/06/2026","vlr":3084.44},{"nf":"123898","dt":"15/06/2026","vlr":3798.0},{"nf":"1262642","dt":"15/06/2026","vlr":667.12}],"venc":[],"emit":null},{"g":"Voke S/ 0353","r":"VOKE S.A.","e":"","c":"Rio de Janeiro","uf":"RJ","qtd":1,"aV":608.22,"v30":0,"v60":0,"v90":0,"v91":0,"total":608.22,"prox":[{"nf":"1277278","dt":"15/06/2026","vlr":608.22}],"venc":[],"emit":null}],"João Ricardo Pereira Cardoso":[{"g":"JAS DO 0187","r":"JAS DO BRASIL AGENCIMENTO LOGISTICO LTDA","e":"alessandra.klayn@jas.com","c":"Rio de Janeiro","uf":"RJ","qtd":2,"aV":928.64,"v30":0,"v60":0,"v90":0,"v91":0,"total":928.64,"prox":[{"nf":"1284877","dt":"25/06/2026","vlr":464.32}],"venc":[],"emit":null},{"g":"PETROBRAS","r":"PETROLEO BRASILEIRO S A PETROBRAS","e":"atendimento.petrobras@brsupply.com.br","c":"Ibirite","uf":"MG","qtd":1462,"aV":884044.07,"v30":136292.52,"v60":6846.38,"v90":4975.3,"v91":70244.97,"total":1102403.24,"prox":[{"nf":"1259323","dt":"15/06/2026","vlr":159.8},{"nf":"1259306","dt":"15/06/2026","vlr":58.47},{"nf":"1259312","dt":"15/06/2026","vlr":59.85}],"venc":[{"nf":"927501","dt":"15/08/2025","vlr":39.57,"d":302},{"nf":"934899","dt":"21/08/2025","vlr":53.26,"d":296},{"nf":"934348","dt":"21/08/2025","vlr":103.04,"d":296},{"nf":"978505","dt":"26/09/2025","vlr":1160.41,"d":260}],"emit":{"total":-3834.62,"itens":[{"ref":"001303755-005","vlr":-58.29,"tipo":"RV","obs":"Transf doc fatmto 91249551 PETROLEO BRASILEIRO S A"},{"ref":"001303738-005","vlr":-58.29,"tipo":"RV","obs":"Transf doc fatmto 91249506 PETROLEO BRASILEIRO S A"},{"ref":"001303701-005","vlr":-54.24,"tipo":"RV","obs":"Transf doc fatmto 91249433 PETROLEO BRASILEIRO S A"}],"alerta":""}},{"g":"ROTA DO OEST","r":"CONCESSIONARIA ROTA DO OESTE SA","e":"jadesonmartins@rotadooeste.com.br","c":"Cuiaba","uf":"MT","qtd":6,"aV":130575.22,"v30":0,"v60":0,"v90":0,"v91":0,"total":130575.22,"prox":[{"nf":"1261233","dt":"15/06/2026","vlr":1580.76},{"nf":"1275255","dt":"29/06/2026","vlr":10985.6},{"nf":"1278735","dt":"02/07/2026","vlr":21734.0}],"venc":[],"emit":null}],"Juliana Regina Urcci Marcal":[{"g":"4 REDES","r":"4 REDES S.A.","e":"atendimento.4redes@brsupply.com.br","c":"Salinas","uf":"MG","qtd":64,"aV":55531.53,"v30":0,"v60":0,"v90":0,"v91":0,"total":55531.53,"prox":[{"nf":"1228128","dt":"25/06/2026","vlr":798.97},{"nf":"1254005","dt":"25/06/2026","vlr":1376.48},{"nf":"1253536","dt":"25/06/2026","vlr":503.95}],"venc":[],"emit":null},{"g":"ADP BR LTDA","r":"ADP BRASIL LTDA","e":"renata.gomes.de.godoi@adp.com","c":"Sao Paulo","uf":"SP","qtd":7,"aV":23528.21,"v30":0,"v60":0,"v90":0,"v91":0,"total":23528.21,"prox":[{"nf":"119620","dt":"25/06/2026","vlr":3680.0},{"nf":"121119","dt":"25/06/2026","vlr":3069.55},{"nf":"127028","dt":"25/06/2026","vlr":2544.84}],"venc":[],"emit":{"total":-6800.31,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"ALSTOM 3254","r":"ALSTOM BRASIL ENERGIA E TRANSPORTE LTDA","e":"alstom.suporte@brsupply.com.br","c":"Taubate","uf":"SP","qtd":94,"aV":255388.06,"v30":0,"v60":0,"v90":0,"v91":0,"total":255388.06,"prox":[{"nf":"102875","dt":"15/06/2026","vlr":6285.0},{"nf":"98032","dt":"15/06/2026","vlr":1123.21},{"nf":"97414","dt":"15/06/2026","vlr":661.69}],"venc":[],"emit":{"total":-4695.89,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"AVERY","r":"AVERY DENNISON DO BRASIL LTDA","e":"alana.rocha@averydennison.com","c":"Vinhedo","uf":"SP","qtd":8,"aV":5051.16,"v30":0,"v60":0,"v90":261.69,"v91":506.52,"total":5819.37,"prox":[{"nf":"112775","dt":"15/06/2026","vlr":887.6},{"nf":"126838","dt":"30/06/2026","vlr":1126.86},{"nf":"126860","dt":"30/06/2026","vlr":358.64}],"venc":[{"nf":"82048","dt":"29/01/2026","vlr":506.52,"d":135},{"nf":"81627","dt":"16/03/2026","vlr":261.69,"d":89}],"emit":{"total":-554.76,"itens":[{"ref":"14052026","vlr":-34.06,"tipo":"DA","obs":"Documento do Cliente 14052026 AVERY DENNISON DO BR"},{"ref":"000082048-007","vlr":-506.52,"tipo":"RV","obs":"Transf doc fatmto 91009338 AVERY DENNISON DO BRASI"},{"ref":"22012026","vlr":-14.18,"tipo":"DA","obs":"Documento do Cliente 22012026 AVERY DENNISON DO BR"}],"alerta":"⚠️ NF 82048 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"CHIESI 0146","r":"CHIESI FARMACEUTICA LTDA","e":"brenda.dourado@chiesi.com","c":"Santana de Parnaiba","uf":"SP","qtd":15,"aV":24394.07,"v30":0,"v60":0,"v90":0,"v91":0,"total":24394.07,"prox":[{"nf":"115594","dt":"22/06/2026","vlr":2058.82},{"nf":"116674","dt":"22/06/2026","vlr":4356.7},{"nf":"126790","dt":"20/07/2026","vlr":756.35}],"venc":[],"emit":{"total":-2104.42,"itens":[{"ref":"10022026","vlr":-121.28,"tipo":"DA","obs":"Documento do Cliente 10022026 SOUZA CRUZ LTDA"},{"ref":"13012026","vlr":-1553.1,"tipo":"DA","obs":"Documento do Cliente 13012026 SOUZA CRUZ LTDA"},{"ref":"11112025","vlr":-430.04,"tipo":"DA","obs":"Documento do Cliente 11112025 SOUZA CRUZ LTDA"}],"alerta":""}},{"g":"CMI BRASIL","r":"CMI BRASIL SERVICOS DE MANUTENCAO DE EQUIPAMENTOS INDUSTRIAIS LTDA","e":"lenilton.inacio@johncockerill.com","c":"Ipatinga","uf":"MG","qtd":5,"aV":3500.07,"v30":0,"v60":0,"v90":0,"v91":0,"total":3500.07,"prox":[{"nf":"1250507","dt":"16/06/2026","vlr":565.72},{"nf":"1250795","dt":"16/06/2026","vlr":940.92},{"nf":"1259920","dt":"29/06/2026","vlr":486.79}],"venc":[],"emit":{"total":-4695.89,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"CONFAB 1919","r":"CONFAB INDUSTRIAL S/A","e":"","c":"Pindamonhangaba","uf":"SP","qtd":11,"aV":12151.97,"v30":0,"v60":0,"v90":0,"v91":0,"total":12151.97,"prox":[{"nf":"119250","dt":"24/06/2026","vlr":774.95},{"nf":"119223","dt":"24/06/2026","vlr":523.15},{"nf":"119071","dt":"24/06/2026","vlr":618.4}],"venc":[],"emit":{"total":-1019.48,"itens":[{"ref":"12022026","vlr":-1019.48,"tipo":"DA","obs":"Documento do Cliente 12022026 ARLANXEO BRASIL S/A"}],"alerta":""}},{"g":"CONSORC 0101","r":"Consorcio De Aluminio Do Maranhao Consorcio Alumar","e":"","c":"Sao Luis","uf":"MA","qtd":24,"aV":40629.86,"v30":0,"v60":0,"v90":0,"v91":0,"total":40629.86,"prox":[{"nf":"1236798","dt":"06/07/2026","vlr":759.9},{"nf":"1253889","dt":"06/07/2026","vlr":1519.8},{"nf":"1245735","dt":"06/07/2026","vlr":3598.0}],"venc":[],"emit":null},{"g":"Compactor","r":"COMPANHIA DE CANETAS COMPACTOR","e":"suprimentos@compactor.com.br","c":"Nova Iguacu","uf":"RJ","qtd":4,"aV":4115.73,"v30":0,"v60":0,"v90":0,"v91":0,"total":4115.73,"prox":[{"nf":"1259905","dt":"29/06/2026","vlr":731.75},{"nf":"1286098","dt":"27/07/2026","vlr":558.11},{"nf":"1305376","dt":"11/08/2026","vlr":2267.76}],"venc":[],"emit":null},{"g":"DANA IN 2100","r":"DANA INDUSTRIAS LTDA","e":"atendimento.danain@brsupply.com.br","c":"Jundiai","uf":"SP","qtd":5,"aV":2439.15,"v30":3671.05,"v60":0,"v90":0,"v91":0,"total":6110.2,"prox":[{"nf":"124318","dt":"06/07/2026","vlr":635.51},{"nf":"134445","dt":"31/07/2026","vlr":663.63},{"nf":"134518","dt":"31/07/2026","vlr":1140.01}],"venc":[{"nf":"110420","dt":"01/06/2026","vlr":1892.99,"d":12},{"nf":"110315","dt":"01/06/2026","vlr":1778.06,"d":12}],"emit":{"total":-2104.42,"itens":[{"ref":"10022026","vlr":-121.28,"tipo":"DA","obs":"Documento do Cliente 10022026 SOUZA CRUZ LTDA"},{"ref":"13012026","vlr":-1553.1,"tipo":"DA","obs":"Documento do Cliente 13012026 SOUZA CRUZ LTDA"},{"ref":"11112025","vlr":-430.04,"tipo":"DA","obs":"Documento do Cliente 11112025 SOUZA CRUZ LTDA"}],"alerta":""}},{"g":"DELPHI","r":"PHINIA DELPHI BRASIL LTDA","e":"eagostini@borgwarner.com","c":"Piracicaba","uf":"SP","qtd":4,"aV":25646.47,"v30":0,"v60":0,"v90":0,"v91":0,"total":25646.47,"prox":[{"nf":"106866","dt":"22/06/2026","vlr":689.28},{"nf":"109549","dt":"29/06/2026","vlr":7608.74},{"nf":"122568","dt":"03/08/2026","vlr":9109.03}],"venc":[],"emit":{"total":-4695.89,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"FUNDACAO SP","r":"FUNDACAO SAO PAULO","e":"ebamador@fundasp.org.br","c":"Sao Paulo","uf":"SP","qtd":66,"aV":12417.64,"v30":0,"v60":0,"v90":0,"v91":0,"total":12417.64,"prox":[{"nf":"127489","dt":"15/06/2026","vlr":60.89},{"nf":"127949","dt":"15/06/2026","vlr":252.19},{"nf":"127820","dt":"15/06/2026","vlr":209.18}],"venc":[],"emit":null},{"g":"Google","r":"GOOGLE BRASIL INTERNET LTDA","e":"pereiraallan@xwf.google.com","c":"Belo Horizonte","uf":"MG","qtd":14,"aV":54408.38,"v30":0,"v60":0,"v90":0,"v91":0,"total":54408.38,"prox":[{"nf":"1202896","dt":"06/07/2026","vlr":3468.34},{"nf":"119451","dt":"23/07/2026","vlr":3601.01},{"nf":"1258921","dt":"28/07/2026","vlr":2298.26}],"venc":[],"emit":{"total":-4695.89,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"ICCAP I 0110","r":"ICCAP IMPLEMENTOS RODOVIARIOS LTDA","e":"compras@iccap.com.br","c":"Sao Luis","uf":"MA","qtd":17,"aV":10776.3,"v30":645.24,"v60":947.14,"v90":0,"v91":0,"total":12368.68,"prox":[{"nf":"1222148","dt":"22/06/2026","vlr":441.17},{"nf":"1222152","dt":"22/06/2026","vlr":1040.42},{"nf":"1222699","dt":"22/06/2026","vlr":705.05}],"venc":[{"nf":"1152320","dt":"15/04/2026","vlr":719.6,"d":59},{"nf":"1153126","dt":"16/04/2026","vlr":227.54,"d":58},{"nf":"1214331","dt":"12/06/2026","vlr":645.24,"d":1}],"emit":null},{"g":"Inspired0190","r":"COLEGIO ELEVA EDUCACAO LTDA","e":"Antonio.Berk@inspirededu.com","c":"Sao Paulo","uf":"SP","qtd":3,"aV":4373.36,"v30":0,"v60":0,"v90":0,"v91":0,"total":4373.36,"prox":[{"nf":"126830","dt":"13/07/2026","vlr":2841.42},{"nf":"135510","dt":"03/08/2026","vlr":487.74},{"nf":"137833","dt":"10/08/2026","vlr":1044.2}],"venc":[],"emit":null},{"g":"KOMATS0178","r":"KOMATSU DO BRASIL LTDA","e":"","c":"Suzano","uf":"SP","qtd":1,"aV":10337.91,"v30":0,"v60":0,"v90":0,"v91":0,"total":10337.91,"prox":[{"nf":"139865","dt":"13/07/2026","vlr":10337.91}],"venc":[],"emit":{"total":-5250.65,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"LOGIN 0124","r":"LOGIN  LOGISTICA INTERMODAL S A","e":"roseane.matos@loginlogistica.com.br","c":"Santos","uf":"SP","qtd":6,"aV":5824.46,"v30":0,"v60":0,"v90":0,"v91":0,"total":5824.46,"prox":[{"nf":"122003","dt":"03/07/2026","vlr":960.96},{"nf":"1278481","dt":"17/07/2026","vlr":539.92},{"nf":"134331","dt":"31/07/2026","vlr":959.04}],"venc":[],"emit":null},{"g":"LYCRA","r":"THE LYCRA COMPANY INDUSTRIA E COMERCIO TEXTIL LTDA","e":"","c":"Paulinia","uf":"SP","qtd":1,"aV":2400.0,"v30":0,"v60":0,"v90":0,"v91":0,"total":2400.0,"prox":[{"nf":"139035","dt":"10/08/2026","vlr":2400.0}],"venc":[],"emit":null},{"g":"Lanxess","r":"ARLANXEO BRASIL SA","e":"","c":"Triunfo","uf":"RS","qtd":38,"aV":52138.49,"v30":0,"v60":0,"v90":0,"v91":0,"total":52138.49,"prox":[{"nf":"692814","dt":"08/07/2026","vlr":2045.95},{"nf":"694079","dt":"09/07/2026","vlr":452.23},{"nf":"703207","dt":"22/07/2026","vlr":1318.77}],"venc":[],"emit":{"total":-4695.89,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"M DIAS BMTZ2","r":"M DIAS BRANCO S/A INDUSTRIA E COMERCIO DE ALIMENTOS","e":"suportemdias@brsupply.com.br","c":"Sao Caetano do Sul","uf":"SP","qtd":73,"aV":191289.75,"v30":1791.7,"v60":0,"v90":0,"v91":0,"total":193081.45,"prox":[{"nf":"1185317","dt":"15/06/2026","vlr":1463.5},{"nf":"1185652","dt":"15/06/2026","vlr":9873.15},{"nf":"1187920","dt":"18/06/2026","vlr":1759.93}],"venc":[{"nf":"1178580","dt":"09/06/2026","vlr":1791.7,"d":4}],"emit":{"total":-111.58,"itens":[{"ref":"15102025","vlr":-111.58,"tipo":"DA","obs":"Documento do Cliente 15102025 M DIAS BRANCO S.A. I"}],"alerta":""}},{"g":"NC EMPR 0165","r":"NEFROCLINICAS HIGIENOPOLIS - SERVICO DE NEFROLOGIA E DIALISE LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":3,"aV":8561.21,"v30":0,"v60":0,"v90":0,"v91":0,"total":8561.21,"prox":[{"nf":"115862","dt":"15/06/2026","vlr":3195.21},{"nf":"133360","dt":"27/07/2026","vlr":2170.79}],"venc":[],"emit":null},{"g":"NEFROCL 0107","r":"NEFROCLINICAS HIGIENOPOLIS - SERVICO DE NEFROLOGIA E DIALISE LTDA","e":"renan.pereira@nefroclinicas.com.br","c":"Sao Paulo","uf":"SP","qtd":4,"aV":4015.01,"v30":0,"v60":0,"v90":0,"v91":0,"total":4015.01,"prox":[{"nf":"123348","dt":"06/07/2026","vlr":1982.36},{"nf":"126596","dt":"13/07/2026","vlr":747.92},{"nf":"133825","dt":"28/07/2026","vlr":624.04}],"venc":[],"emit":null},{"g":"PURCOM 0103","r":"PURCOM QUIMICA LTDA","e":"b.almeida@purcom.com.br","c":"Barueri","uf":"SP","qtd":2,"aV":4090.81,"v30":0,"v60":0,"v90":0,"v91":0,"total":4090.81,"prox":[{"nf":"121998","dt":"18/06/2026","vlr":2853.41},{"nf":"137631","dt":"24/07/2026","vlr":1237.4}],"venc":[],"emit":{"total":-2104.42,"itens":[{"ref":"10022026","vlr":-121.28,"tipo":"DA","obs":"Documento do Cliente 10022026 SOUZA CRUZ LTDA"},{"ref":"13012026","vlr":-1553.1,"tipo":"DA","obs":"Documento do Cliente 13012026 SOUZA CRUZ LTDA"},{"ref":"11112025","vlr":-430.04,"tipo":"DA","obs":"Documento do Cliente 11112025 SOUZA CRUZ LTDA"}],"alerta":""}},{"g":"SCARFAC 0185","r":"SCARFACE INDUSTRIA E COMERCIO DE CONFECCOES LTDA","e":"compras.brasilia@ricardoalmeida.com.br","c":"Brasilia","uf":"DF","qtd":8,"aV":4797.53,"v30":2283.4,"v60":0,"v90":0,"v91":0,"total":7080.93,"prox":[{"nf":"1245033","dt":"15/06/2026","vlr":1093.51},{"nf":"1224349","dt":"22/06/2026","vlr":640.93},{"nf":"1242278","dt":"09/07/2026","vlr":812.69}],"venc":[{"nf":"1244145","dt":"12/06/2026","vlr":1141.7,"d":1}],"emit":null},{"g":"SKF DO BRAS","r":"SKF DO BRASIL LTDA","e":"suporte.skf@brsupply.com.br","c":"Cajamar","uf":"SP","qtd":33,"aV":41707.49,"v30":0,"v60":0,"v90":0,"v91":0,"total":41707.49,"prox":[{"nf":"116942","dt":"17/06/2026","vlr":6070.5},{"nf":"121118","dt":"01/07/2026","vlr":273.86},{"nf":"121480","dt":"01/07/2026","vlr":118.71}],"venc":[],"emit":{"total":-5250.65,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"SMART","r":"ZILIA TECHNOLOGIES INDUSTRIA E COMERCIO DE COMPONENTES ELETRONICOS LTDA","e":"Alexsandro.Lima@ziliatech.com","c":"Manaus","uf":"AM","qtd":16,"aV":15030.37,"v30":0,"v60":0,"v90":0,"v91":0,"total":15030.37,"prox":[{"nf":"1221839","dt":"19/06/2026","vlr":287.14},{"nf":"1221900","dt":"19/06/2026","vlr":956.05},{"nf":"1221901","dt":"19/06/2026","vlr":954.92}],"venc":[],"emit":null},{"g":"SOUZA CRUZ","r":"SOUZA CRUZ LTDA","e":"","c":"Goiania","uf":"GO","qtd":8,"aV":11114.12,"v30":1483.03,"v60":0,"v90":0,"v91":0,"total":12597.15,"prox":[{"nf":"1232377","dt":"30/06/2026","vlr":793.72},{"nf":"1266128","dt":"04/08/2026","vlr":2346.64},{"nf":"1279982","dt":"17/08/2026","vlr":1744.3}],"venc":[{"nf":"1194513","dt":"26/05/2026","vlr":1483.03,"d":18}],"emit":{"total":-2104.42,"itens":[{"ref":"10022026","vlr":-121.28,"tipo":"DA","obs":"Documento do Cliente 10022026 SOUZA CRUZ LTDA"},{"ref":"13012026","vlr":-1553.1,"tipo":"DA","obs":"Documento do Cliente 13012026 SOUZA CRUZ LTDA"},{"ref":"11112025","vlr":-430.04,"tipo":"DA","obs":"Documento do Cliente 11112025 SOUZA CRUZ LTDA"}],"alerta":""}},{"g":"TECLASER","r":"INSTITUTO BRASILEIRO DE OFTALMOLOGIAS.A. - IBOL","e":"compras@ibol.com.br","c":"Rio de Janeiro","uf":"RJ","qtd":7,"aV":6235.0,"v30":0,"v60":0,"v90":0,"v91":0,"total":6235.0,"prox":[{"nf":"702590","dt":"22/06/2026","vlr":1047.64},{"nf":"1264738","dt":"06/07/2026","vlr":994.22},{"nf":"1255684","dt":"24/07/2026","vlr":735.64}],"venc":[],"emit":null},{"g":"TORRE ALTA","r":"MAHLE RODO DIESEL COMERCIO DE COMBUSTIVEIS LTDA","e":"adm29@postosmahle.com.br","c":"Sao Simao","uf":"GO","qtd":191,"aV":182819.74,"v30":726.1,"v60":0,"v90":0,"v91":0,"total":183545.84,"prox":[{"nf":"1223905","dt":"22/06/2026","vlr":795.0},{"nf":"1224374","dt":"22/06/2026","vlr":766.76},{"nf":"1223693","dt":"22/06/2026","vlr":717.69}],"venc":[{"nf":"662859","dt":"27/05/2026","vlr":726.1,"d":17}],"emit":null},{"g":"UNIMED BR","r":"UNIMED DO BRASIL CONFEDERACAO NAC DAS COOPERATIVAS MED","e":"comprasbrasil@unimed.coop.br","c":"Sao Paulo","uf":"SP","qtd":1,"aV":399.74,"v30":0,"v60":0,"v90":0,"v91":0,"total":399.74,"prox":[{"nf":"139036","dt":"27/07/2026","vlr":399.74}],"venc":[],"emit":{"total":-5250.65,"itens":[{"ref":"01062026","vlr":-865.26,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"01062026","vlr":-2729.28,"tipo":"DA","obs":"Documento do Cliente 1062026 ALSTOM BRASIL ENERGIA"},{"ref":"16032026","vlr":-81.87,"tipo":"DA","obs":"Documento do Cliente 16032026 ALSTOM BRASIL ENERGI"}],"alerta":""}},{"g":"XERIUMPIRACI","r":"ANDRITZ FABRICS AND ROLLS INDUSTRIA ECOMERCIO S.A.","e":"amanda.colombo@andritz.com","c":"Piracicaba","uf":"SP","qtd":28,"aV":80150.3,"v30":0,"v60":0,"v90":0,"v91":0,"total":80150.3,"prox":[{"nf":"1247190","dt":"15/06/2026","vlr":1000.4},{"nf":"116003","dt":"16/06/2026","vlr":7835.05},{"nf":"120651","dt":"16/06/2026","vlr":772.47}],"venc":[],"emit":null}],"Leonardo Vinicius Lino da Silva":[{"g":"ABENGOA 0147","r":"COX CONSTRUCAO BRASIL LTDA.","e":"luana.teixeira@coxabengoa.com","c":"Colinas do Tocantins","uf":"TO","qtd":6,"aV":71770.91,"v30":6345.11,"v60":0,"v90":0,"v91":0,"total":78116.02,"prox":[{"nf":"1250293","dt":"18/06/2026","vlr":3108.73},{"nf":"1237952","dt":"30/06/2026","vlr":59335.99}],"venc":[{"nf":"1250852","dt":"18/05/2026","vlr":6345.11,"d":26}],"emit":null},{"g":"AGILENT 0606","r":"AGILENT TECHNOLOGIES BRASIL LTDA","e":"marcia.souza@non.agilent.com","c":"Barueri","uf":"SP","qtd":4,"aV":8607.39,"v30":0,"v60":0,"v90":0,"v91":0,"total":8607.39,"prox":[{"nf":"125665","dt":"25/06/2026","vlr":320.92},{"nf":"127918","dt":"29/06/2026","vlr":6526.42},{"nf":"133548","dt":"13/07/2026","vlr":1439.13}],"venc":[],"emit":null},{"g":"ATACADA 0101","r":"ATACADAO DO PAPEL LTDA","e":"","c":"Camacari","uf":"BA","qtd":4,"aV":41671.35,"v30":0,"v60":0,"v90":0,"v91":0,"total":41671.35,"prox":[{"nf":"4011","dt":"15/06/2026","vlr":7461.79},{"nf":"3980","dt":"29/06/2026","vlr":14844.39},{"nf":"4144","dt":"09/07/2026","vlr":11903.17}],"venc":[],"emit":null},{"g":"ATACADA 0124","r":"ATACADAO PAPELEX LTDA","e":"","c":"Rio de Janeiro","uf":"RJ","qtd":5,"aV":0,"v30":0,"v60":0,"v90":0,"v91":115470.66,"total":115470.66,"prox":[],"venc":[{"nf":"1826","dt":"08/10/2024","vlr":25111.04,"d":613},{"nf":"1850","dt":"21/10/2024","vlr":24897.1,"d":600},{"nf":"533587","dt":"21/10/2024","vlr":15541.02,"d":600},{"nf":"539048","dt":"28/10/2024","vlr":17721.6,"d":593}],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"ATT PER 0106","r":"COMERCIAL ATT LTDA","e":"lojasantaluzia@attboticario.com.br","c":"Santa Luzia","uf":"MG","qtd":23,"aV":17658.43,"v30":760.83,"v60":0,"v90":0,"v91":0,"total":18419.26,"prox":[{"nf":"1278145","dt":"17/06/2026","vlr":972.7},{"nf":"1284273","dt":"24/06/2026","vlr":763.0},{"nf":"1291557","dt":"01/07/2026","vlr":545.61}],"venc":[{"nf":"1294536","dt":"03/06/2026","vlr":760.83,"d":10}],"emit":{"total":-1348.95,"itens":[{"ref":"001294536-005","vlr":-760.83,"tipo":"RV","obs":"Transf doc fatmto 91233836 ATT PERFUMARIA E COSMET"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":"⚠️ NF 1294536 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"BRAZAO","r":"BRAZAO COMERCIAL LTDA","e":"","c":"Santos","uf":"SP","qtd":5,"aV":97277.9,"v30":0,"v60":0,"v90":0,"v91":0,"total":97277.9,"prox":[{"nf":"4045","dt":"29/06/2026","vlr":9729.45},{"nf":"4130","dt":"01/07/2026","vlr":25939.31}],"venc":[],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"BRENNTAG","r":"BRENNTAG QUIMICA BRASIL LTDA.","e":"emerson.Silva@brenntag.com","c":"Guarulhos","uf":"SP","qtd":30,"aV":108695.87,"v30":0,"v60":0,"v90":0,"v91":0,"total":108695.87,"prox":[{"nf":"116947","dt":"16/06/2026","vlr":9267.04},{"nf":"1251104","dt":"19/06/2026","vlr":395.58},{"nf":"1253308","dt":"22/06/2026","vlr":3019.25}],"venc":[],"emit":null},{"g":"Braile","r":"DISTRIBUIDORA DE PAPEIS BRAILE LTDA","e":"","c":"Porto Alegre","uf":"RS","qtd":3,"aV":308771.93,"v30":0,"v60":0,"v90":0,"v91":0,"total":308771.93,"prox":[{"nf":"4039","dt":"29/06/2026","vlr":45329.75},{"nf":"4043","dt":"29/06/2026","vlr":33904.23},{"nf":"4126","dt":"27/07/2026","vlr":229537.95}],"venc":[],"emit":null},{"g":"CASA DA 0177","r":"CASA DAS IMPRESSORAS LTDA","e":"","c":"Belo Horizonte","uf":"MG","qtd":4,"aV":32741.17,"v30":0,"v60":0,"v90":0,"v91":0,"total":32741.17,"prox":[{"nf":"4097","dt":"18/06/2026","vlr":7118.42},{"nf":"3984","dt":"22/06/2026","vlr":11385.69}],"venc":[],"emit":null},{"g":"COLOR RJ0155","r":"COLOR MIDIA MIX SUPRIMENTOS INFORMATICA LTDA","e":"","c":"Niteroi","uf":"RJ","qtd":2,"aV":8607.81,"v30":0,"v60":0,"v90":0,"v91":0,"total":8607.81,"prox":[{"nf":"4122","dt":"25/06/2026","vlr":1854.56},{"nf":"4141","dt":"08/07/2026","vlr":6753.25}],"venc":[],"emit":{"total":-2587.74,"itens":[{"ref":"18112025","vlr":-2587.74,"tipo":"DA","obs":"Documento do Cliente 18112025 COLOR MIDIA SUPRIMEN"}],"alerta":""}},{"g":"CREATIV 0316","r":"CREATIVE COPIAS LTDA","e":"","c":"Serra","uf":"ES","qtd":5,"aV":65478.66,"v30":0,"v60":0,"v90":0,"v91":0,"total":65478.66,"prox":[{"nf":"4108","dt":"18/06/2026","vlr":1534.76},{"nf":"4109","dt":"22/06/2026","vlr":17471.35},{"nf":"4000","dt":"07/07/2026","vlr":11529.34}],"venc":[],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"CREATIV0002","r":"CREATIVE COPIAS LTDA","e":"","c":"Serra","uf":"ES","qtd":1,"aV":11326.26,"v30":0,"v60":0,"v90":0,"v91":0,"total":11326.26,"prox":[{"nf":"3960","dt":"22/06/2026","vlr":11326.26}],"venc":[],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"DAIKIN 0115","r":"DAIKIN AR CONDICIONADO BRASIL LTDA","e":"renata.monteiro@daikin.com.br","c":"Sao Paulo","uf":"SP","qtd":3,"aV":2061.27,"v30":0,"v60":0,"v90":0,"v91":0,"total":2061.27,"prox":[{"nf":"123200","dt":"19/06/2026","vlr":516.1},{"nf":"132065","dt":"10/07/2026","vlr":1029.07}],"venc":[],"emit":null},{"g":"DIGITAL 0105","r":"DIGITAL GOLD 333 ELETRONICOS E INFORMATICA LTDA","e":"","c":"Rio de Janeiro","uf":"RJ","qtd":4,"aV":4144.73,"v30":0,"v60":0,"v90":120.38,"v91":0,"total":4265.11,"prox":[{"nf":"4072","dt":"22/06/2026","vlr":1816.32},{"nf":"4120","dt":"02/07/2026","vlr":1550.34},{"nf":"4133","dt":"09/07/2026","vlr":778.07}],"venc":[{"nf":"3953","dt":"19/03/2026","vlr":120.38,"d":86}],"emit":{"total":-120.38,"itens":[{"ref":"000003953-006","vlr":-120.38,"tipo":"RV","obs":"Transf doc fatmto 91098613 DIGITAL GOLD 333 ELETRO"}],"alerta":"⚠️ NF 3953 aparece VENCIDA na carteira — conciliar antes de cobrar!"}},{"g":"Epl Bra 0135","r":"EPL BRASIL LTDA.","e":"","c":"Seropedica","uf":"RJ","qtd":5,"aV":0,"v30":0,"v60":0,"v90":0,"v91":37110.25,"total":37110.25,"prox":[],"venc":[{"nf":"1132140","dt":"20/02/2026","vlr":7422.05,"d":113}],"emit":null},{"g":"F J P S 0118","r":"F J P SUPRIMENTOS PARA INFORMATICA LTDA","e":"","c":"Londrina","uf":"PR","qtd":1,"aV":998.97,"v30":0,"v60":0,"v90":0,"v91":0,"total":998.97,"prox":[{"nf":"4082","dt":"13/07/2026","vlr":998.97}],"venc":[],"emit":null},{"g":"F J P S 0207","r":"F J P SUPRIMENTOS PARA INFORMATICA LTDA","e":"","c":"Londrina","uf":"PR","qtd":9,"aV":15816.11,"v30":0,"v60":0,"v90":0,"v91":0,"total":15816.11,"prox":[{"nf":"4014","dt":"16/06/2026","vlr":1704.49},{"nf":"4028","dt":"23/06/2026","vlr":2266.99},{"nf":"4117","dt":"25/06/2026","vlr":3064.06}],"venc":[],"emit":null},{"g":"FABESUL","r":"FABESUL COMERCIO DE SUPRIMENTOS LTDA","e":"","c":"Porto Alegre","uf":"RS","qtd":4,"aV":53897.65,"v30":0,"v60":0,"v90":0,"v91":0,"total":53897.65,"prox":[{"nf":"4069","dt":"22/06/2026","vlr":14579.1},{"nf":"4112","dt":"10/07/2026","vlr":37169.22},{"nf":"4123","dt":"13/07/2026","vlr":1526.75}],"venc":[],"emit":null},{"g":"FERNAND 0149","r":"OBERO COMERCIO DE SUPRIMENTOS DE INFORMATICA LTDA","e":"","c":"Nilopolis","uf":"RJ","qtd":8,"aV":856301.84,"v30":0,"v60":0,"v90":0,"v91":0,"total":856301.84,"prox":[{"nf":"3965","dt":"24/06/2026","vlr":94942.27},{"nf":"1287812","dt":"25/06/2026","vlr":2678.4},{"nf":"4054","dt":"29/06/2026","vlr":255295.78}],"venc":[],"emit":null},{"g":"GECORE","r":"GECORE COMERCIAL DISTRIBUIDORA LTDA","e":"","c":"Vitoria","uf":"ES","qtd":10,"aV":20384.8,"v30":0,"v60":0,"v90":0,"v91":0,"total":20384.8,"prox":[{"nf":"4024","dt":"19/06/2026","vlr":4173.39},{"nf":"1252843","dt":"22/06/2026","vlr":734.74},{"nf":"1232015","dt":"30/06/2026","vlr":376.0}],"venc":[],"emit":null},{"g":"HOTEL ACCOR","r":"ATRIO HOTEIS S.A.","e":"atendimento.astore@brsupply.com.br","c":"Rio de Janeiro","uf":"RJ","qtd":114,"aV":128371.15,"v30":14032.04,"v60":0,"v90":3035.89,"v91":684.87,"total":146123.95,"prox":[{"nf":"1278705","dt":"15/06/2026","vlr":911.85},{"nf":"129252","dt":"16/06/2026","vlr":870.66},{"nf":"129447","dt":"16/06/2026","vlr":1076.47}],"venc":[{"nf":"1132988","dt":"25/02/2026","vlr":228.29,"d":108},{"nf":"1119969","dt":"31/03/2026","vlr":542.63,"d":74},{"nf":"1212492","dt":"09/04/2026","vlr":1408.0,"d":65},{"nf":"119566","dt":"22/05/2026","vlr":1683.56,"d":22}],"emit":{"total":-456.96,"itens":[{"ref":"14042026","vlr":-456.96,"tipo":"DA","obs":"Documento do Cliente 14042026 ATRIO HOTEIS S.A."}],"alerta":""}},{"g":"INNO IN 0177","r":"INNO INFORMATICA LTDA","e":"","c":"Criciuma","uf":"SC","qtd":12,"aV":27430.54,"v30":5028.69,"v60":0,"v90":0,"v91":0,"total":32459.23,"prox":[{"nf":"3958","dt":"22/06/2026","vlr":2678.46},{"nf":"1254994","dt":"23/06/2026","vlr":4299.62},{"nf":"4042","dt":"29/06/2026","vlr":919.88}],"venc":[{"nf":"4083","dt":"12/06/2026","vlr":5028.69,"d":1}],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"KOPELL","r":"LEPOK DISTRIBUICAO E LOGISTICA LTDA","e":"","c":"Sao Paulo","uf":"SP","qtd":16,"aV":139188.25,"v30":0,"v60":0,"v90":0,"v91":0,"total":139188.25,"prox":[{"nf":"3964","dt":"15/06/2026","vlr":29338.92},{"nf":"4031","dt":"22/06/2026","vlr":8659.98},{"nf":"4105","dt":"14/07/2026","vlr":5301.01}],"venc":[],"emit":{"total":-1341.08,"itens":[{"ref":"001294536-005","vlr":-760.83,"tipo":"RV","obs":"Transf doc fatmto 91233836 ATT PERFUMARIA E COSMET"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"}],"alerta":""}},{"g":"LC MO RJ0106","r":"L C MOREIRA SUPRIMENTOS DE INFORMATICA LTDA","e":"","c":"Volta Redonda","uf":"RJ","qtd":6,"aV":9285.78,"v30":0,"v60":0,"v90":0,"v91":0,"total":9285.78,"prox":[{"nf":"4075","dt":"15/06/2026","vlr":1347.51},{"nf":"4096","dt":"15/06/2026","vlr":1248.87},{"nf":"4029","dt":"23/06/2026","vlr":845.76}],"venc":[],"emit":null},{"g":"LJ INFO","r":"GIGABIT SUPRIMENTOS COMERCIO E IMPORTACAO LTDA - ME","e":"","c":"Sao Paulo","uf":"SP","qtd":6,"aV":65507.34,"v30":0,"v60":0,"v90":0,"v91":0,"total":65507.34,"prox":[{"nf":"4128","dt":"29/06/2026","vlr":11797.65},{"nf":"4131","dt":"02/07/2026","vlr":10037.92}],"venc":[],"emit":{"total":-2587.74,"itens":[{"ref":"18112025","vlr":-2587.74,"tipo":"DA","obs":"Documento do Cliente 18112025 COLOR MIDIA SUPRIMEN"}],"alerta":""}},{"g":"LMZ COM 0120","r":"LMZ COMERCIO DE ARTIGOS PARA ESCRITORIO LTDA","e":"","c":"Campo Grande","uf":"MS","qtd":5,"aV":1094.94,"v30":0,"v60":0,"v90":0,"v91":0,"total":1094.94,"prox":[{"nf":"737207","dt":"06/07/2026","vlr":239.63},{"nf":"694659","dt":"09/07/2026","vlr":136.44}],"venc":[],"emit":null},{"g":"LOGIN I 0515","r":"LOGIN INFORMATICA COMERCIO E REPRESENTACAO LTDA","e":"","c":"Salvador","uf":"BA","qtd":1,"aV":46508.49,"v30":0,"v60":0,"v90":0,"v91":0,"total":46508.49,"prox":[{"nf":"3974","dt":"24/06/2026","vlr":46508.49}],"venc":[],"emit":null},{"g":"MICROL","r":"MICROL INFORMATICA LTDA","e":"","c":"Curitiba","uf":"PR","qtd":23,"aV":28250.36,"v30":1242.6,"v60":0,"v90":0,"v91":0,"total":29492.96,"prox":[{"nf":"703400","dt":"23/06/2026","vlr":964.11},{"nf":"683563","dt":"24/06/2026","vlr":2841.4},{"nf":"3975","dt":"25/06/2026","vlr":2952.86}],"venc":[{"nf":"718112","dt":"12/06/2026","vlr":1242.6,"d":1}],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"MICROL  0263","r":"MICROL INFORMATICA LTDA","e":"","c":"Joinville","uf":"SC","qtd":1,"aV":858.02,"v30":0,"v60":0,"v90":0,"v91":0,"total":858.02,"prox":[{"nf":"1234323","dt":"01/07/2026","vlr":858.02}],"venc":[],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"MICROS  0106","r":"MICROS E ACESSORIOS INFORMATICA LTDA","e":"","c":"Belo Horizonte","uf":"MG","qtd":4,"aV":2939.91,"v30":3545.0,"v60":0,"v90":0,"v91":0,"total":6484.91,"prox":[{"nf":"4078","dt":"13/07/2026","vlr":1469.94}],"venc":[{"nf":"3928","dt":"04/06/2026","vlr":2075.06,"d":9},{"nf":"4078","dt":"11/06/2026","vlr":1469.94,"d":2}],"emit":{"total":-1341.08,"itens":[{"ref":"001294536-005","vlr":-760.83,"tipo":"RV","obs":"Transf doc fatmto 91233836 ATT PERFUMARIA E COSMET"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"}],"alerta":""}},{"g":"MIRANDRN0580","r":"MIRANDA COMPUTACAO E COMERCIO LTDA","e":"","c":"Natal","uf":"RN","qtd":6,"aV":77386.18,"v30":0,"v60":0,"v90":0,"v91":0,"total":77386.18,"prox":[{"nf":"3943","dt":"15/06/2026","vlr":2552.46},{"nf":"3956","dt":"18/06/2026","vlr":5987.8},{"nf":"4129","dt":"29/06/2026","vlr":20151.86}],"venc":[],"emit":{"total":-1341.08,"itens":[{"ref":"001294536-005","vlr":-760.83,"tipo":"RV","obs":"Transf doc fatmto 91233836 ATT PERFUMARIA E COSMET"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"}],"alerta":""}},{"g":"MORADA DO SO","r":"RODOVIARIO MORADA DO SOL LTDA (EM RECUPERACAO JUDICIAL)","e":"leonardo.ferreira@morada.com.br","c":"Americana","uf":"SP","qtd":1,"aV":0,"v30":0,"v60":0,"v90":0,"v91":1549.74,"total":1549.74,"prox":[],"venc":[{"nf":"508854","dt":"19/09/2024","vlr":1549.74,"d":632}],"emit":null},{"g":"Maquila 0725","r":"MAQUILAR COM DE MAQUINAS P ESC E ASSISTENCIA TEC LTDA","e":"","c":"Fortaleza","uf":"CE","qtd":3,"aV":8548.62,"v30":0,"v60":0,"v90":0,"v91":0,"total":8548.62,"prox":[{"nf":"3961","dt":"22/06/2026","vlr":3068.32},{"nf":"4040","dt":"29/06/2026","vlr":2740.15}],"venc":[],"emit":null},{"g":"PARK ECOM","r":"PARQUE ARTIGOS PARA INFORMATICA E SERVICOS LTDA","e":"","c":"Rio de Janeiro","uf":"RJ","qtd":5,"aV":290259.86,"v30":0,"v60":0,"v90":0,"v91":0,"total":290259.86,"prox":[{"nf":"4093","dt":"15/06/2026","vlr":46891.23},{"nf":"4052","dt":"29/06/2026","vlr":74791.25}],"venc":[],"emit":null},{"g":"Primex  0187","r":"PRIMEX DISTRIBUIDORA DE TECNOLOGIA LTDA","e":"","c":"Sao Jose do Rio Preto","uf":"SP","qtd":1,"aV":4222.38,"v30":0,"v60":0,"v90":0,"v91":0,"total":4222.38,"prox":[{"nf":"3968","dt":"24/06/2026","vlr":4222.38}],"venc":[],"emit":null},{"g":"SODECIA","r":"SODECIA AUTOMOTIVE SAO PAULO LTDA.","e":"atendimento.sodecia@brsupply.com.br","c":"Aracariguama","uf":"SP","qtd":2,"aV":2003.71,"v30":0,"v60":0,"v90":0,"v91":0,"total":2003.71,"prox":[{"nf":"137132","dt":"08/07/2026","vlr":1670.47},{"nf":"137942","dt":"10/07/2026","vlr":333.24}],"venc":[],"emit":null},{"g":"SUPRY POINT","r":"SUPRY POINT INFORMATICA LTDA","e":"","c":"Campinas","uf":"SP","qtd":12,"aV":88241.25,"v30":12849.87,"v60":0,"v90":0,"v91":0,"total":101091.12,"prox":[{"nf":"4005","dt":"15/06/2026","vlr":5998.94},{"nf":"3949","dt":"16/06/2026","vlr":9806.86},{"nf":"4016","dt":"19/06/2026","vlr":4289.96}],"venc":[{"nf":"4057","dt":"11/06/2026","vlr":12849.87,"d":2}],"emit":null},{"g":"SYMA CO 0136","r":"SYMA COMPUTADORES LTDA","e":"","c":"Maringa","uf":"PR","qtd":21,"aV":33704.47,"v30":2405.84,"v60":0,"v90":0,"v91":0,"total":36110.31,"prox":[{"nf":"3948","dt":"16/06/2026","vlr":2981.78},{"nf":"4102","dt":"19/06/2026","vlr":1213.52},{"nf":"4027","dt":"23/06/2026","vlr":1673.36}],"venc":[{"nf":"4084","dt":"12/06/2026","vlr":2405.84,"d":1}],"emit":{"total":-588.12,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}},{"g":"TECNO I 0457","r":"TECNO INDUSTRIA E COMERCIO DE COMPUTADORES LTDA","e":"","c":"Fortaleza","uf":"CE","qtd":2,"aV":7969.08,"v30":0,"v60":0,"v90":0,"v91":0,"total":7969.08,"prox":[{"nf":"4061","dt":"03/07/2026","vlr":3984.48}],"venc":[],"emit":{"total":-1341.08,"itens":[{"ref":"001294536-005","vlr":-760.83,"tipo":"RV","obs":"Transf doc fatmto 91233836 ATT PERFUMARIA E COSMET"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"},{"ref":"001173384-1","vlr":-128.44,"tipo":"RV","obs":"Transf doc fatmto 001173384-1 LEPOK DISTRIBUICAO E"}],"alerta":""}},{"g":"VERACEL","r":"VERACEL CELULOSE S.A.","e":"robson.aquino@veracel.com.br","c":"Eunapolis","uf":"BA","qtd":2,"aV":6233.46,"v30":0,"v60":0,"v90":0,"v91":0,"total":6233.46,"prox":[{"nf":"1275874","dt":"15/06/2026","vlr":2498.6},{"nf":"1282509","dt":"22/06/2026","vlr":3734.86}],"venc":[],"emit":{"total":-456.96,"itens":[{"ref":"14042026","vlr":-456.96,"tipo":"DA","obs":"Documento do Cliente 14042026 ATRIO HOTEIS S.A."}],"alerta":""}},{"g":"VITORIA 0603","r":"VITORIA HOTEIS LTDA","e":"compras.paulinia@vitoriahoteis.com.br","c":"Paulinia","uf":"SP","qtd":1,"aV":460.74,"v30":0,"v60":0,"v90":0,"v91":0,"total":460.74,"prox":[{"nf":"130552","dt":"18/06/2026","vlr":460.74}],"venc":[],"emit":{"total":-1045.08,"itens":[{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.39,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"},{"ref":"001116976-001","vlr":-172.4,"tipo":"RV","obs":"Transf doc fatmto 001116976-001 SYMA COMPUTADORES"}],"alerta":""}}]};

const EXECS = Object.keys(CANAL);
const CORES = {
  "Danilo Artioli":                {"cor":"#1565C0","bg":"#E3F2FD","ini":"DA"},
  "Evelyn Rodrigues de Souza":     {"cor":"#6A1B9A","bg":"#F3E5F5","ini":"ER"},
  "Gilvania Pitanga Soares":       {"cor":"#2E7D32","bg":"#E8F5E9","ini":"GP"},
  "Jessica Antunes da Silva":      {"cor":"#C62828","bg":"#FFEBEE","ini":"JA"},
  "João Ricardo Pereira Cardoso":  {"cor":"#E65100","bg":"#FFF3E0","ini":"JR"},
  "Juliana Regina Urcci Marcal":   {"cor":"#00838F","bg":"#E0F7FA","ini":"JU"},
  "Leonardo Vinicius Lino da Silva":{"cor":"#4527A0","bg":"#EDE7F6","ini":"LV"},
};

const R = v => Math.abs(v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const TODAY = new Date(2026,5,13);
const diffD = s => { const[d,m,a]=s.split("/").map(Number); return Math.round((new Date(a,m-1,d)-TODAY)/864e5); };

const statusOf = c => {
  const vT = c.v30+c.v60+c.v90+c.v91;
  if(c.v91>0) return {l:"Crítico",    cor:"#C62828",bg:"#FFEBEE",s:4};
  if(c.v90>0) return {l:"Alto Risco", cor:"#D84315",bg:"#FBE9E7",s:3};
  if(c.v60>0) return {l:"Atenção",    cor:"#EF6C00",bg:"#FFF3E0",s:2};
  if(c.v30>0) return {l:"Vencendo",   cor:"#F9A825",bg:"#FFFDE7",s:1};
  if(c.emit)  return {l:"Crédito",    cor:"#6A1B9A",bg:"#F3E5F5",s:0};
  return             {l:"Normal",     cor:"#2E7D32",bg:"#E8F5E9",s:0};
};


const STATUS = {
  protestado:  {l:"Protestado",                    e:"🔴", cor:"#B71C1C", bg:"#FFEBEE"},
  conciliar:   {l:"A Conciliar",                   e:"💜", cor:"#6A1B9A", bg:"#F3E5F5"},
  cobrado:     {l:"Cobrado",                       e:"📧", cor:"#1565C0", bg:"#E3F2FD"},
  pago:        {l:"Regularizado / Pago",            e:"✅", cor:"#1B5E20", bg:"#E8F5E9"},
  aguardando:  {l:"Aguardando Retorno",             e:"⏳", cor:"#E65100", bg:"#FFF3E0"},
  negociacao:  {l:"Em Negociação",                  e:"🔵", cor:"#006064", bg:"#E0F7FA"},
  comprovante: {l:"Comprovante de Entrega Enviado", e:"📦", cor:"#4A148C", bg:"#EDE7F6"},
  prorrogacao: {l:"Prorrogação Solicitada",         e:"📅", cor:"#827717", bg:"#F9FBE7"},
  nfd:         {l:"Ocorrência Ag. NFD",             e:"📋", cor:"#37474F", bg:"#ECEFF1"},
};
const STATUS_LIST = Object.entries(STATUS);

// ── Storage helpers (localStorage para Vercel) ──────────────────────────────
const localStorageGet = async (key) => {
  try { const v = localStorage.getItem(key); return v ? { value: v } : null; } catch(e) { return null; }
};
const localStorageSet = async (key, value) => {
  try { localStorage.setItem(key, value); } catch(e) {}
};

const MSGS = {
  preventivo: c=>`Prezado(a) ${c.r},\n\nEntramos em contato para informar que identificamos títulos a vencer em breve em sua conta.\n\n`+(c.prox.length?`📋 Próximos vencimentos:\n${c.prox.map(p=>`  • NF ${p.nf}  |  Venc. ${p.dt}  |  ${R(p.vlr)}`).join("\n")}\n\n`:``)+`Solicitamos confirmação de recebimento das notas.\n\nAtenciosamente,\nEquipe Financeira`,
  vencido:    c=>`Prezado(a) ${c.r},\n\n⚠️ Títulos VENCIDOS em sua conta:\n\n`+c.venc.map(v=>`  • NF ${v.nf}  |  ${v.dt}  |  ${R(v.vlr)}  |  ${v.d} dias em atraso`).join("\n")+`\n\nSolicitamos regularização urgente.\n\nAtenciosamente,\nEquipe Financeira`,
  lancamento: c=>`Prezado(a) ${c.r},\n\nLembrete: as notas abaixo precisam ser lançadas em seu sistema:\n\n`+(c.prox.length?c.prox.map(p=>`  📄 NF ${p.nf}  |  Venc. ${p.dt}  |  ${R(p.vlr)}`).join("\n"):"verificar pendências")+`\n\nO não lançamento pode causar atraso no pagamento.\n\nAtenciosamente,\nEquipe Financeira`,
  conciliacao:c=>`Prezado(a) ${c.r},\n\nIdentificamos crédito pendente de conciliação: ${R(c.emit?.total||0)}\n\n`+(c.emit?.itens||[]).map(i=>`  • Ref. ${i.ref}  |  ${R(i.vlr)}  |  ${i.tipo}\n    ${i.obs}`).join("\n")+`\n\nConciliar antes de efetuar qualquer pagamento.\n\nAtenciosamente,\nEquipe Financeira`,

};

const ASSUNTOS = {
  preventivo:  c=>`BRSupply | Títulos a vencer — ${c.r}`,
  vencido:     c=>`BRSupply | Títulos vencidos — ${c.r}`,
  lancamento:  c=>`BRSupply | Lançamento de NF pendente — ${c.r}`,
  conciliacao: c=>`BRSupply | Crédito a conciliar — ${c.r}`,
};

const openEmail = (email, from, c, tipo) => {
  const to  = encodeURIComponent(email || "");
  const sub = encodeURIComponent(ASSUNTOS[tipo](c));
  // Prepend sender instruction if 'from' is set, so user knows which inbox to open
  const fromNote = from ? `⚠️ ENVIAR DE: ${from}\n${"─".repeat(40)}\n\n` : "";
  const body = encodeURIComponent(fromNote + MSGS[tipo](c));
  window.open(`mailto:${to}?subject=${sub}&body=${body}`, "_blank");
};

export default function App() {
  // ── dados dinâmicos (começa com seed, pode ser atualizado via upload) ──
  const [canalData, setCanalData]   = useState(CANAL);
  const [baseDate, setBaseDate]     = useState("12/06/2026");
  const [showImport, setShowImport] = useState(false);
  const [importing, setImporting]   = useState(false);
  const [importMsg, setImportMsg]   = useState(null); // {tipo:"ok"|"erro", txt}
  const [dragOver, setDragOver]     = useState(null); // "titulos"|"emit"
  const [fileTitulos, setFileTitulos] = useState(null);
  const [fileEmit, setFileEmit]       = useState(null);

  const [execSel, setExecSel]     = useState(null);
  const [clientSel, setClientSel] = useState(null);
  const [search, setSearch]       = useState("");
  const [filtSt, setFiltSt]       = useState("todos");
  const [msgT, setMsgT]           = useState("preventivo");
  const [copied, setCopied]       = useState(false);
  const [log, setLog]             = useState([]);
  const [subTab, setSubTab]       = useState("clientes");
  // Para: overrides { "EXEC|GRUPO": "email@cliente.com" }
  const [emailMap, setEmailMap]     = useState({});
  // De: (remetente) { "EXEC|GRUPO": "atendimento.xxx@brsupply.com.br" }
  const [fromMap, setFromMap]       = useState({});
  const [editingEmail, setEditingEmail] = useState(null); // "para|EXEC|GRUPO" or "de|EXEC|GRUPO"
  const [emailDraft, setEmailDraft]     = useState("");
  // NF status: { "EXEC|GRUPO|NF": "protestado"|"cobrado"|... }
  const [nfStatus, setNfStatus]         = useState({});

  const emailKey = (exec, grupo) => `${exec}|${grupo}`;
  const getEmail = (exec, c) => emailMap[emailKey(exec, c.g)] ?? c.e;

  // Load Para: and De: from storage on mount
  useState(() => {
    (async () => {
      try {
        const r1 = localStorageGet("emails_canal_para");
        if (r1?.value) setEmailMap(JSON.parse(r1.value));
        const r2 = localStorageGet("emails_canal_de");
        if (r2?.value) setFromMap(JSON.parse(r2.value));
      } catch(e) {}
    })();
  });

  const getFrom  = (exec, c) => fromMap[emailKey(exec, c.g)] ?? "";

  const saveEmailPara = async (exec, grupo) => {
    const newMap = {...emailMap, [emailKey(exec, grupo)]: emailDraft.trim()};
    setEmailMap(newMap);
    try { localStorageSet("emails_canal_para", JSON.stringify(newMap)); } catch(e) {}
    setLog(prev => [{id:Date.now(), exec, g:grupo, acao:`E-mail Para: → ${emailDraft.trim()||"(removido)"}`, ts:new Date().toLocaleString("pt-BR")}, ...prev.slice(0,99)]);
    setEditingEmail(null);
  };

  const saveEmailDe = async (exec, grupo) => {
    const newMap = {...fromMap, [emailKey(exec, grupo)]: emailDraft.trim()};
    setFromMap(newMap);
    try { localStorageSet("emails_canal_de", JSON.stringify(newMap)); } catch(e) {}
    setLog(prev => [{id:Date.now(), exec, g:grupo, acao:`E-mail De: → ${emailDraft.trim()||"(removido)"}`, ts:new Date().toLocaleString("pt-BR")}, ...prev.slice(0,99)]);
    setEditingEmail(null);
  };

  // Unified save dispatcher
  const saveEmail = async (exec, grupo) => {
    if (editingEmail?.startsWith("de|")) await saveEmailDe(exec, grupo);
    else await saveEmailPara(exec, grupo);
  };

  // Load NF statuses from storage on mount
  useState(()=>{
    (async()=>{
      try{ const r=localStorageGet("nf_status_canal"); if(r?.value) setNfStatus(JSON.parse(r.value)); }catch(e){}
    })();
  });

  const nfKey = (exec, grupo, nf) => `${exec}|${grupo}|${nf}`;
  const getNfStatus = (exec, grupo, nf) => nfStatus[nfKey(exec, grupo, nf)] || null;
  const setNfStatusVal = async (exec, grupo, nf, val) => {
    const key = nfKey(exec, grupo, nf);
    const newMap = val ? {...nfStatus, [key]: val} : Object.fromEntries(Object.entries(nfStatus).filter(([k])=>k!==key));
    setNfStatus(newMap);
    try{ localStorageSet("nf_status_canal", JSON.stringify(newMap)); }catch(e){}
    const st = val ? STATUS[val] : null;
    setLog(prev=>[{id:Date.now(), exec, g:grupo, acao:`NF ${nf}: ${st ? st.e+" "+st.l : "status removido"}`, ts:new Date().toLocaleString("pt-BR")}, ...prev.slice(0,99)]);
  };

  // Enrich all clients
  const EXECS_DYN = Object.keys(canalData);

  const allClients = useMemo(() => {
    const out = {};
    EXECS_DYN.forEach(ex => {
      out[ex] = (canalData[ex]||[]).map(c => {
        const total = c.aV+c.v30+c.v60+c.v90+c.v91;
        const vencidoT = c.v30+c.v60+c.v90+c.v91;
        const liquidoC = c.emit ? total+c.emit.total : total;
        return {...c, total, vencidoT, liquidoC, st: statusOf(c)};
      }).sort((a,b)=>b.st.s-a.st.s||b.vencidoT-a.vencidoT);
    });
    return out;
  }, []);

  // Canal-level totals
  const canalTotals = useMemo(() => {
    let t={total:0,aV:0,v30:0,v60:0,v90:0,v91:0,emit:0,clientes:0,criticos:0,alertas:0,protestadoCount:0,protestadoValor:0};
    EXECS_DYN.forEach(ex => {
      (allClients[ex]||[]).forEach(c => {
        t.total+=c.total; t.aV+=c.aV; t.v30+=c.v30; t.v60+=c.v60; t.v90+=c.v90; t.v91+=c.v91;
        t.clientes++;
        if(c.emit) t.emit+=c.emit.total;
        if(c.st.s>=4) t.criticos++;
        c.prox.forEach(p=>{ if(diffD(p.dt)>=0&&diffD(p.dt)<=7) t.alertas++; });
        // count protestado NFs
        [...c.venc,...c.prox].forEach(nf=>{
          if(getNfStatus(ex,c.g,nf.nf)==="protestado"){ t.protestadoCount++; t.protestadoValor+=nf.vlr; }
        });
      });
    });
    return t;
  }, [allClients, nfStatus]);

  // Exec summaries
  const execSummaries = useMemo(() => {
    return EXECS_DYN.map(ex => {
      const cls = allClients[ex]||[];
      // Protestado: NFs with status "protestado" for this exec
      const protestadoNFs = [];
      cls.forEach(c=>[...c.venc,...c.prox].forEach(nf=>{
        if(getNfStatus(ex,c.g,nf.nf)==="protestado") protestadoNFs.push({vlr:nf.vlr,g:c.g});
      }));
      return {
        nome: ex,
        clientes: cls.length,
        total: cls.reduce((s,c)=>s+c.total,0),
        vencido: cls.reduce((s,c)=>s+c.vencidoT,0),
        v91: cls.reduce((s,c)=>s+c.v91,0),
        emit: Math.abs(cls.reduce((s,c)=>s+(c.emit?.total||0),0)),
        criticos: cls.filter(c=>c.st.s>=4).length,
        comEmit: cls.filter(c=>c.emit).length,
        alertas: cls.reduce((s,c)=>s+c.prox.filter(p=>diffD(p.dt)>=0&&diffD(p.dt)<=7).length,0),
        protestadoCount: protestadoNFs.length,
        protestadoValor: Math.round(protestadoNFs.reduce((s,n)=>s+n.vlr,0)*100)/100,
      };
    }).sort((a,b)=>b.vencido-a.vencido);
  }, [allClients]);

  // Current exec clients
  const execClients = useMemo(() => {
    if(!execSel) return [];
    let list = allClients[execSel];
    if(search) list = list.filter(c=>c.g.toLowerCase().includes(search.toLowerCase())||c.r.toLowerCase().includes(search.toLowerCase()));
    if(filtSt==="crit")   list = list.filter(c=>c.st.s>=4);
    if(filtSt==="risco")  list = list.filter(c=>c.st.s===3);
    if(filtSt==="atenc")  list = list.filter(c=>c.st.s>=1&&c.st.s<=2);
    if(filtSt==="emit")   list = list.filter(c=>!!c.emit);
    if(filtSt==="normal") list = list.filter(c=>c.st.s===0&&!c.emit);
    return list;
  }, [execSel, search, filtSt, allClients]);

  const execAlerts = useMemo(() => {
    if(!execSel) return [];
    const items=[];
    (allClients[execSel]||[]).forEach(c=>c.prox.forEach(p=>{const d=diffD(p.dt);if(d>=0&&d<=7)items.push({...p,g:c.g,r:c.r,cid:c.g,d,hasEmit:!!c.emit,cRef:c});}));
    return items.sort((a,b)=>a.d-b.d);
  },[execSel,allClients]);

  const copyMsg = c => {
    navigator.clipboard.writeText(MSGS[msgT](c)).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2200);});
    setLog(p=>[{id:Date.now(),exec:execSel,g:c.g,acao:msgT,ts:new Date().toLocaleString("pt-BR")},...p.slice(0,99)]);
  };

  const ini = name => CORES[name]?.ini || name.slice(0,2).toUpperCase();
  const cor = name => CORES[name]?.cor || "#607D8B";
  const bg  = name => CORES[name]?.bg  || "#ECEFF1";

  // ── PROCESS UPLOAD ─────────────────────────────────────────────────────
  const processUpload = async () => {
    if (!fileTitulos || !fileEmit) {
      setImportMsg({tipo:"erro", txt:"Selecione os dois arquivos antes de importar."});
      return;
    }
    setImporting(true);
    setImportMsg(null);
    try {
      // ── Parse CSV de títulos ──
      const csvText = await fileTitulos.text();
      const parsed = Papa.parse(csvText, {header:true, delimiter:";", encoding:"utf-8", skipEmptyLines:true});
      const rows = parsed.data;

      // ── Parse XLSX do EMIT ──
      const emitBuf = await fileEmit.arrayBuffer();
      const wb = XLSX.read(emitBuf, {type:"array"});
      const emitRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

      // ── Compute today ──
      const today = new Date();
      const todayStr = today.toLocaleDateString("pt-BR");

      const parseBR = (s) => {
        if(!s) return null;
        const [d,m,a] = String(s).split("/").map(Number);
        return new Date(a, m-1, d);
      };
      const diffDays = (dt) => {
        if(!dt) return 0;
        return Math.round((today - dt) / 86400000);
      };
      const bucket = (d) => {
        if(d <= 0)   return "aVencer";
        if(d <= 30)  return "ate30";
        if(d <= 60)  return "d30_60";
        if(d <= 90)  return "d61_90";
        return "acima91";
      };
      const parseVlr = (s) => parseFloat(String(s||"0").replace(",",".")) || 0;

      // ── Build EMIT credit map: exec → { nomeCliente → [{ref,vlr,tipo,obs}] } ──
      const emitMap = {};
      emitRows.forEach(r => {
        const exec = String(r["EXECUTIVO"]||"").trim();
        const vlr  = parseFloat(r["Montante (ME)"]) || 0;
        if(vlr >= 0) return; // only credits (negative)
        const nome = String(r["Nome do cliente"]||"").trim();
        if(!emitMap[exec]) emitMap[exec] = {};
        if(!emitMap[exec][nome]) emitMap[exec][nome] = [];
        emitMap[exec][nome].push({
          ref:  String(r["Referência"]||"").trim(),
          vlr:  Math.round(vlr * 100)/100,
          tipo: String(r["Tipo lçto.contábil"]||"").trim(),
          obs:  String(r["Texto de item"]||"").trim().slice(0,80),
        });
      });

      // ── Group titles by ExecVendas → NmCliente ──
      const byExec = {};
      rows.forEach(r => {
        const exec   = String(r["ExecVendas"]||"").trim();
        const client = String(r["NmCliente"]||"").trim();
        const vlr    = parseVlr(r["VlrSaldo"]);
        if(vlr <= 0) return;
        if(!byExec[exec]) byExec[exec] = {};
        if(!byExec[exec][client]) byExec[exec][client] = [];
        byExec[exec][client].push(r);
      });

      const newCanal = {};
      Object.entries(byExec).forEach(([exec, clients]) => {
        newCanal[exec] = [];
        const execEmitCredits = emitMap[exec] || {};

        Object.entries(clients).forEach(([nm, cRows]) => {
          const emails = [...new Set(cRows.map(r=>String(r["email"]||"").trim()).filter(Boolean))];
          const razao  = String(cRows[0]["RazaoSocial"]||nm).trim();
          const cidade = String(cRows[0]["Cidade"]||"").trim();
          const uf     = String(cRows[0]["DestUF"]||"").trim();
          const qtd    = cRows.length;

          const bkts = {aVencer:0,ate30:0,d30_60:0,d61_90:0,acima91:0};
          cRows.forEach(r => {
            const dt = parseBR(r["DtVencimento"]);
            const d  = diffDays(dt);
            const b  = bucket(d);
            bkts[b] += parseVlr(r["VlrSaldo"]);
          });
          Object.keys(bkts).forEach(k => { bkts[k] = Math.round(bkts[k]*100)/100; });

          // proximos (a vencer) — top 5 by date
          const aVencerRows = cRows.filter(r => {
            const d = diffDays(parseBR(r["DtVencimento"]));
            return d <= 0;
          }).sort((a,b) => parseBR(a["DtVencimento"]) - parseBR(b["DtVencimento"]));
          const seen = new Set();
          const prox = [];
          aVencerRows.forEach(r => {
            const nf = String(r["NrNotaFiscal"]||"").trim();
            if(seen.has(nf)) return; seen.add(nf);
            prox.push({nf, dt: r["DtVencimento"], vlr: Math.round(parseVlr(r["VlrSaldo"])*100)/100});
            if(prox.length >= 5) return;
          });

          // vencidos — top 6 by oldest
          const vencRows = cRows.filter(r => diffDays(parseBR(r["DtVencimento"])) > 0)
            .sort((a,b) => parseBR(a["DtVencimento"]) - parseBR(b["DtVencimento"]));
          const seenV = new Set();
          const venc = [];
          vencRows.forEach(r => {
            const nf = String(r["NrNotaFiscal"]||"").trim();
            if(seenV.has(nf)) return; seenV.add(nf);
            const d = diffDays(parseBR(r["DtVencimento"]));
            venc.push({nf, dt: r["DtVencimento"], vlr: Math.round(parseVlr(r["VlrSaldo"])*100)/100, d});
            if(venc.length >= 6) return;
          });

          // EMIT credits match
          const razaoUp = razao.toUpperCase();
          let matchedCredits = [];
          let matchedTotal = 0;
          Object.entries(execEmitCredits).forEach(([emitNome, itens]) => {
            const words1 = new Set(razaoUp.split(/\s+/).slice(0,3));
            const words2 = new Set(emitNome.toUpperCase().split(/\s+/).slice(0,3));
            const common = [...words1].filter(w => words2.has(w));
            if(common.length > 0) {
              matchedCredits.push(...itens);
              matchedTotal += itens.reduce((s,i) => s+i.vlr, 0);
            }
          });
          let emit = null;
          if(matchedCredits.length > 0) {
            emit = {total: Math.round(matchedTotal*100)/100, itens: matchedCredits.slice(0,6)};
            const vencNfs = new Set(venc.map(v=>v.nf));
            for(const it of matchedCredits) {
              const refNf = it.ref.split("-")[0].replace(/^0+/,"");
              if(vencNfs.has(refNf)) {
                emit.alerta = `⚠️ NF ${refNf} aparece VENCIDA na carteira — conciliar antes de cobrar!`;
                break;
              }
            }
          }

          newCanal[exec].push({
            g: nm, r: razao, e: emails[0]||"", c: cidade, uf, qtd,
            aV: bkts.aVencer, v30: bkts.ate30, v60: bkts.d30_60,
            v90: bkts.d61_90, v91: bkts.acima91,
            total: Math.round((bkts.aVencer+bkts.ate30+bkts.d30_60+bkts.d61_90+bkts.acima91)*100)/100,
            prox, venc, emit,
          });
        });

        // sort by worst status
        newCanal[exec].sort((a,b) => {
          const score = c => c.v91>0?4:c.v90>0?3:c.v60>0?2:c.v30>0?1:0;
          return score(b)-score(a) || b.v91+b.v90+b.v60+b.v30 - (a.v91+a.v90+a.v60+a.v30);
        });
      });

      // ── Totals for subtitle ──
      let totalTitulos = 0, totalClientes = 0;
      Object.values(newCanal).forEach(cl => { totalClientes += cl.length; cl.forEach(c => { totalTitulos += c.qtd; }); });

      setCanalData(newCanal);
      setBaseDate(todayStr);
      setExecSel(null);
      setClientSel(null);
      setImportMsg({tipo:"ok", txt:`✅ Base atualizada! ${totalTitulos.toLocaleString()} títulos · ${totalClientes} clientes · ${Object.keys(newCanal).length} executivos`});
      setFileTitulos(null);
      setFileEmit(null);
      setTimeout(()=>{ setShowImport(false); setImportMsg(null); }, 3000);
    } catch(err) {
      console.error(err);
      setImportMsg({tipo:"erro", txt:`❌ Erro ao processar: ${err.message}`});
    }
    setImporting(false);
  };

  // ── CANAL OVERVIEW (no exec selected)
  if(!execSel) return (
    <div style={{fontFamily:"'Inter','Segoe UI',sans-serif",background:"#EEF2F7",minHeight:"100vh"}}>
      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#1A237E 0%,#283593 50%,#3949AB 100%)",padding:"18px 24px",boxShadow:"0 3px 16px rgba(26,35,126,.4)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{color:"#C5CAE9",fontSize:10,fontWeight:700,letterSpacing:2.5,textTransform:"uppercase",marginBottom:3}}>Canal Acordos · Visão Gerencial</div>
        <div style={{color:"#fff",fontSize:20,fontWeight:800}}>Painel de Cobrança & Prevenção</div>
        <div style={{color:"#9FA8DA",fontSize:11,marginTop:2}}>{Object.values(canalData).reduce((s,cl)=>s+cl.reduce((a,c)=>a+c.qtd,0),0).toLocaleString()} títulos · {Object.values(canalData).reduce((s,cl)=>s+cl.length,0)} clientes · {Object.keys(canalData).length} executivos · Base: {baseDate}</div>
      </div>
      <button onClick={()=>setShowImport(true)} style={{background:"rgba(255,255,255,.15)",border:"1.5px solid rgba(255,255,255,.35)",borderRadius:10,padding:"10px 18px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:7,whiteSpace:"nowrap",flexShrink:0}}>📥 Atualizar Base</button>
      {/* ── IMPORT MODAL ── */}
      {showImport&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget)setShowImport(false);}}>
          <div style={{background:"#fff",borderRadius:16,padding:32,width:560,maxWidth:"95vw",boxShadow:"0 12px 50px rgba(0,0,0,.25)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
              <div>
                <div style={{fontWeight:800,fontSize:18,color:"#1A2332"}}>📥 Atualizar Base</div>
                <div style={{fontSize:12,color:"#607D8B",marginTop:3}}>Arraste ou selecione os dois arquivos exportados do sistema</div>
              </div>
              <button onClick={()=>{setShowImport(false);setImportMsg(null);}} style={{background:"#ECEFF1",border:"none",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,color:"#607D8B",fontWeight:600}}>✕ Fechar</button>
            </div>

            {/* Drop zones */}
            {[
              {key:"titulos", label:"Base de Títulos", sub:"CSV com ponto-e-vírgula (;) — mesmo formato exportado do sistema", accept:".csv,.txt", icon:"📋", color:"#1565C0", bg:"#E3F2FD", file:fileTitulos, setFile:setFileTitulos},
              {key:"emit",    label:"BASE EMIT (SAP)", sub:"Excel .xlsx exportado do SAP UI5",                                   accept:".xlsx",     icon:"💜", color:"#6A1B9A", bg:"#F3E5F5", file:fileEmit,    setFile:setFileEmit},
            ].map(z=>(
              <div key={z.key}
                onDragOver={e=>{e.preventDefault();setDragOver(z.key);}}
                onDragLeave={()=>setDragOver(null)}
                onDrop={e=>{e.preventDefault();setDragOver(null);const f=e.dataTransfer.files[0];if(f)z.setFile(f);}}
                style={{border:`2px dashed ${dragOver===z.key?z.color:"#CFD8DC"}`,borderRadius:12,padding:"20px 22px",marginBottom:14,background:dragOver===z.key?z.bg:z.file?"#F8FFF8":"#FAFAFA",transition:"all .2s",cursor:"pointer"}}
                onClick={()=>document.getElementById(`inp-${z.key}`).click()}>
                <input id={`inp-${z.key}`} type="file" accept={z.accept} style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(f)z.setFile(f);e.target.value="";}}/>
                <div style={{display:"flex",alignItems:"center",gap:14}}>
                  <span style={{fontSize:32}}>{z.icon}</span>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,fontSize:14,color:"#1A2332"}}>{z.label}</div>
                    <div style={{fontSize:11,color:"#90A4AE",marginTop:2}}>{z.sub}</div>
                  </div>
                  {z.file?(
                    <div style={{textAlign:"right"}}>
                      <div style={{fontSize:12,fontWeight:700,color:"#2E7D32"}}>✅ {z.file.name}</div>
                      <div style={{fontSize:10,color:"#90A4AE"}}>{(z.file.size/1024).toFixed(0)} KB</div>
                      <button onClick={e=>{e.stopPropagation();z.setFile(null);}} style={{fontSize:10,color:"#C62828",background:"none",border:"none",cursor:"pointer",marginTop:2}}>remover</button>
                    </div>
                  ):(
                    <div style={{fontSize:12,color:"#B0BEC5",fontStyle:"italic"}}>arraste aqui<br/>ou clique para selecionar</div>
                  )}
                </div>
              </div>
            ))}

            {/* Status message */}
            {importMsg&&(
              <div style={{padding:"10px 14px",borderRadius:9,marginBottom:14,background:importMsg.tipo==="ok"?"#E8F5E9":"#FFEBEE",color:importMsg.tipo==="ok"?"#1B5E20":"#B71C1C",fontWeight:600,fontSize:13}}>
                {importMsg.txt}
              </div>
            )}

            {/* Action buttons */}
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>{setShowImport(false);setImportMsg(null);setFileTitulos(null);setFileEmit(null);}}
                style={{flex:1,padding:12,borderRadius:10,border:"1.5px solid #CFD8DC",background:"#fff",cursor:"pointer",fontWeight:600,fontSize:14,color:"#607D8B"}}>
                Cancelar
              </button>
              <button
                onClick={processUpload}
                disabled={!fileTitulos||!fileEmit||importing}
                style={{flex:2,padding:12,borderRadius:10,border:"none",cursor:fileTitulos&&fileEmit&&!importing?"pointer":"not-allowed",fontWeight:700,fontSize:14,background:fileTitulos&&fileEmit&&!importing?"#1565C0":"#B0BEC5",color:"#fff",transition:"background .2s"}}>
                {importing?"⏳ Processando...":"📥 Importar e Atualizar Base"}
              </button>
            </div>
            <div style={{fontSize:11,color:"#90A4AE",marginTop:10,textAlign:"center"}}>
              ✅ Status, e-mails editados e histórico permanecem salvos após a atualização
            </div>
          </div>
        </div>
      )}

      <div style={{padding:"20px 24px"}}>
        {/* Canal KPIs */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:14,marginBottom:20}}>
          {[
            {l:"Total Carteira",  v:R(canalTotals.total),   i:"💰",c:"#1565C0",sub:`${canalTotals.clientes} clientes`},
            {l:"Vencido +91d",    v:R(canalTotals.v91),     i:"🚨",c:"#C62828",sub:`${canalTotals.criticos} críticos`},
            {l:"Vencido até 90d", v:R(canalTotals.v30+canalTotals.v60+canalTotals.v90),i:"⚠️",c:"#E65100",sub:"em risco"},
            {l:"Créditos EMIT",   v:R(Math.abs(canalTotals.emit)),i:"💜",c:"#6A1B9A",sub:"a conciliar"},
            {l:"🔴 Protestado",   v:R(canalTotals.protestadoValor),i:"🔴",c:"#B71C1C",sub:`${canalTotals.protestadoCount} NFs`,highlight:canalTotals.protestadoCount>0},
            {l:"NFs Venc. 7 dias",v:canalTotals.alertas,   i:"📅",c:"#00838F",sub:"alertas ativos"},
          ].map(k=>(
            <div key={k.l} style={{background:k.highlight?"#FFEBEE":"#fff",borderRadius:12,padding:"18px 20px",boxShadow:"0 1px 6px rgba(0,0,0,.08)",borderTop:`4px solid ${k.c}`,transition:"background .3s"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontSize:11,color:k.highlight?"#B71C1C":"#607D8B",fontWeight:k.highlight?700:500,marginBottom:5}}>{k.l}</div>
                  <div style={{fontSize:18,fontWeight:800,color:k.c}}>{k.v}</div>
                  <div style={{fontSize:11,color:k.highlight?"#C62828":"#90A4AE",marginTop:3}}>{k.sub}</div>
                </div>
                <span style={{fontSize:26}}>{k.i}</span>
              </div>
            </div>
          ))}
        </div>

        {/* Aging canal */}
        <div style={{background:"#fff",borderRadius:12,padding:22,boxShadow:"0 1px 6px rgba(0,0,0,.08)",marginBottom:20}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:14}}>Aging — Canal Acordos (5.103 títulos)</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12}}>
            {[{l:"A Vencer",v:canalTotals.aV,c:"#1976D2"},{l:"Até 30d",v:canalTotals.v30,c:"#F9A825"},{l:"31–60d",v:canalTotals.v60,c:"#EF6C00"},{l:"61–90d",v:canalTotals.v90,c:"#D84315"},{l:"+91d",v:canalTotals.v91,c:"#C62828"}].map(ag=>(
              <div key={ag.l} style={{textAlign:"center"}}>
                <div style={{height:8,background:"#ECEFF1",borderRadius:4,overflow:"hidden",marginBottom:8}}>
                  <div style={{height:"100%",width:`${(ag.v/canalTotals.total)*100}%`,background:ag.c,borderRadius:4}}/>
                </div>
                <div style={{fontWeight:800,color:ag.c,fontSize:15}}>{R(ag.v)}</div>
                <div style={{fontSize:11,color:"#607D8B"}}>{ag.l}</div>
                <div style={{fontSize:10,color:"#B0BEC5"}}>{((ag.v/canalTotals.total)*100).toFixed(1)}%</div>
              </div>
            ))}
          </div>
        </div>

        {/* Executivos cards */}
        <div style={{fontWeight:700,fontSize:15,marginBottom:14,color:"#1A2332"}}>Selecionar Executivo</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:12,marginBottom:20}}>
          {execSummaries.map(ex=>(
            <div key={ex.nome} onClick={()=>{setExecSel(ex.nome);setClientSel(null);setSearch("");setFiltSt("todos");setSubTab("clientes");}} style={{background:"#fff",borderRadius:12,padding:"16px 14px",boxShadow:"0 1px 6px rgba(0,0,0,.08)",cursor:"pointer",borderTop:`4px solid ${cor(ex.nome)}`,transition:"transform .15s,box-shadow .15s"}} onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-3px)";e.currentTarget.style.boxShadow="0 6px 20px rgba(0,0,0,.15)";}} onMouseLeave={e=>{e.currentTarget.style.transform="none";e.currentTarget.style.boxShadow="0 1px 6px rgba(0,0,0,.08)";}}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
                <div style={{width:38,height:38,borderRadius:"50%",background:bg(ex.nome),display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:13,color:cor(ex.nome),flexShrink:0}}>{ini(ex.nome)}</div>
                <div style={{fontSize:11,fontWeight:700,color:"#1A2332",lineHeight:1.3}}>{ex.nome.split(" ").slice(0,2).join(" ")}</div>
              </div>
              <div style={{marginBottom:8}}>
                <div style={{fontSize:10,color:"#90A4AE"}}>Carteira</div>
                <div style={{fontWeight:800,color:cor(ex.nome),fontSize:14}}>{R(ex.total)}</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                <div style={{background:"#FFF5F5",borderRadius:6,padding:"5px 7px"}}>
                  <div style={{fontSize:9,color:"#90A4AE"}}>Vencido</div>
                  <div style={{fontWeight:700,color:"#C62828",fontSize:12}}>{R(ex.vencido)}</div>
                </div>
                <div style={{background:"#F3E5F5",borderRadius:6,padding:"5px 7px"}}>
                  <div style={{fontSize:9,color:"#90A4AE"}}>EMIT</div>
                  <div style={{fontWeight:700,color:"#6A1B9A",fontSize:12}}>{R(ex.emit)}</div>
                </div>
              </div>
              <div style={{display:"flex",gap:6,marginTop:8,flexWrap:"wrap"}}>
                {ex.criticos>0&&<span style={{fontSize:9,background:"#FFEBEE",color:"#C62828",fontWeight:700,padding:"2px 6px",borderRadius:10}}>🚨 {ex.criticos}</span>}
                {ex.alertas>0&&<span style={{fontSize:9,background:"#E0F7FA",color:"#00838F",fontWeight:700,padding:"2px 6px",borderRadius:10}}>📅 {ex.alertas}</span>}
                <span style={{fontSize:9,background:"#F5F5F5",color:"#607D8B",fontWeight:600,padding:"2px 6px",borderRadius:10}}>{ex.clientes} cli.</span>
              </div>
            </div>
          ))}
        </div>

        {/* Ranking vencidos */}
        <div style={{background:"#fff",borderRadius:12,padding:22,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:14}}>🏆 Ranking por Vencido no Canal</div>
          {execSummaries.map((ex,i)=>(
            <div key={ex.nome} onClick={()=>{setExecSel(ex.nome);setClientSel(null);}} style={{display:"flex",alignItems:"center",gap:14,padding:"10px 8px",borderRadius:8,cursor:"pointer",marginBottom:4}} onMouseEnter={e=>e.currentTarget.style.background="#F5F7FA"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <div style={{width:26,height:26,borderRadius:"50%",background:i<3?"#FFF3E0":"#F5F5F5",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,color:i===0?"#E65100":i===1?"#F9A825":i===2?"#8D6E63":"#90A4AE",flexShrink:0}}>{i+1}</div>
              <div style={{width:34,height:34,borderRadius:"50%",background:bg(ex.nome),display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,color:cor(ex.nome),flexShrink:0}}>{ini(ex.nome)}</div>
              <div style={{flex:1}}>
                <div style={{fontWeight:600,fontSize:13}}>{ex.nome}</div>
                <div style={{height:5,background:"#ECEFF1",borderRadius:3,marginTop:4,overflow:"hidden"}}>
                  <div style={{height:"100%",width:`${(ex.vencido/execSummaries[0].vencido)*100}%`,background:cor(ex.nome),borderRadius:3}}/>
                </div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontWeight:800,color:cor(ex.nome),fontSize:14}}>{R(ex.vencido)}</div>
                <div style={{fontSize:10,color:"#90A4AE"}}>{ex.clientes} clientes</div>
              </div>
              {ex.criticos>0&&<span style={{background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,padding:"3px 8px",borderRadius:20}}>🚨 {ex.criticos}</span>}
            </div>
          ))}
        </div>
      </div>
    </div>
  );

  // ── EXEC VIEW (exec selected)
  const execInfo = execSummaries.find(e=>e.nome===execSel);
  return (
    <div style={{fontFamily:"'Inter','Segoe UI',sans-serif",background:"#EEF2F7",minHeight:"100vh"}}>
      {/* Header exec */}
      <div style={{background:`linear-gradient(135deg,${cor(execSel)} 0%,${cor(execSel)}cc 100%)`,padding:"16px 22px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,boxShadow:"0 3px 14px rgba(0,0,0,.25)"}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <button onClick={()=>{setExecSel(null);setClientSel(null);}} style={{background:"rgba(255,255,255,.2)",border:"none",borderRadius:8,padding:"6px 12px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:13}}>← Canal</button>
          <div style={{width:42,height:42,borderRadius:"50%",background:"rgba(255,255,255,.25)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:15,color:"#fff"}}>{ini(execSel)}</div>
          <div>
            <div style={{color:"rgba(255,255,255,.7)",fontSize:10,fontWeight:600,letterSpacing:2,textTransform:"uppercase"}}>Executivo</div>
            <div style={{color:"#fff",fontSize:17,fontWeight:800}}>{execSel.split(" ").slice(0,3).join(" ")}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {["clientes","alertas","conciliacao","mensagens","status","historico"].map(t=>(
            <button key={t} onClick={()=>setSubTab(t)} style={{padding:"7px 13px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:600,fontSize:12,background:subTab===t?"#fff":"rgba(255,255,255,.18)",color:subTab===t?cor(execSel):"#fff"}}>
              {t==="clientes"?"👥 Clientes":t==="alertas"?"🔔 Alertas"+(execAlerts.length>0?` (${execAlerts.length})`:""):t==="conciliacao"?"💜 EMIT":t==="mensagens"?"💬 Mensagens":t==="status"?"🏷️ Status":"📋 Histórico"}
            </button>
          ))}
        </div>
      </div>

      {/* Exec KPIs */}
      <div style={{background:"#fff",borderBottom:"1px solid #ECEFF1",padding:"14px 22px",display:"flex",gap:20,overflowX:"auto"}}>
        {[
          {l:"Carteira",    v:R(execInfo.total),              c:cor(execSel)},
          {l:"Vencido",     v:R(execInfo.vencido),            c:"#C62828"},
          {l:"+91 dias",    v:R(execInfo.v91),                c:"#C62828"},
          {l:"Créditos EMIT",v:R(execInfo.emit),             c:"#6A1B9A"},
          {l:"🔴 Protestado",v:execInfo.protestadoCount>0?`${execInfo.protestadoCount} NF${execInfo.protestadoCount>1?"s":""} · ${R(execInfo.protestadoValor)}`:"—", c:"#B71C1C", highlight:execInfo.protestadoCount>0},
          {l:"Clientes",    v:execInfo.clientes,              c:cor(execSel)},
          {l:"Críticos",    v:execInfo.criticos,              c:"#C62828"},
          {l:"Com EMIT",    v:execInfo.comEmit,               c:"#6A1B9A"},
          {l:"Alertas 7d",  v:execInfo.alertas,               c:"#00838F"},
        ].map(k=>(
          <div key={k.l} style={{textAlign:"center",minWidth:80,background:k.highlight?"#FFEBEE":"transparent",borderRadius:k.highlight?8:0,padding:k.highlight?"6px 10px":0}}>
            <div style={{fontSize:10,color:"#90A4AE",marginBottom:2}}>{k.l}</div>
            <div style={{fontWeight:800,color:k.c,fontSize:16,whiteSpace:"nowrap"}}>{k.v}</div>
          </div>
        ))}
      </div>

      <div style={{padding:"18px 22px"}}>

      {/* CLIENTES */}
      {subTab==="clientes"&&<>
        <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar cliente..." style={{flex:1,minWidth:200,padding:"9px 14px",borderRadius:10,border:"1.5px solid #CFD8DC",fontSize:14,outline:"none",background:"#fff"}}/>
          {[["todos","Todos"],["crit","🚨 Crítico"],["risco","⚠️ Risco"],["atenc","🟡 Atenção"],["emit","💜 EMIT"],["normal","✅ Normal"]].map(([k,l])=>(
            <button key={k} onClick={()=>setFiltSt(k)} style={{padding:"8px 12px",borderRadius:8,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:filtSt===k?cor(execSel):"#E3F2FD",color:filtSt===k?"#fff":"#1565C0"}}>{l}</button>
          ))}
          <span style={{fontSize:13,color:"#607D8B"}}>{execClients.length}</span>
        </div>
        <div style={{overflowX:"auto",background:"#fff",borderRadius:12,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:cor(execSel),color:"#fff"}}>
              {["Cliente","UF","NFs","A Vencer","Vencido","+91d","Crédito EMIT","Status",""].map(h=>(
                <th key={h} style={{padding:"11px 10px",textAlign:h==="Cliente"||h==="UF"||h===""?"left":"right",whiteSpace:"nowrap",fontWeight:600,fontSize:11}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>{execClients.map((c,i)=>(
              <tr key={c.g+i} style={{background:c.emit?i%2===0?"#FAF5FF":"#F3E5F5":i%2===0?"#fff":"#F8FAFC",cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background="#EDE7F6"} onMouseLeave={e=>e.currentTarget.style.background=c.emit?i%2===0?"#FAF5FF":"#F3E5F5":i%2===0?"#fff":"#F8FAFC"} onClick={()=>{setClientSel(c);setSubTab("mensagens");}}>
                <td style={{padding:"9px 10px",fontWeight:600,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  <div>{c.g}</div><div style={{fontSize:10,color:"#90A4AE",fontWeight:400}}>{c.c}</div>
                </td>
                <td style={{padding:"9px 10px",fontSize:11,color:"#607D8B"}}>{c.uf}</td>
                <td style={{padding:"9px 10px",textAlign:"right",color:"#607D8B"}}>{c.qtd}</td>
                <td style={{padding:"9px 10px",textAlign:"right",color:"#1565C0",fontSize:12}}>{R(c.aV)}</td>
                <td style={{padding:"9px 10px",textAlign:"right",color:c.vencidoT>0?"#C62828":"#CFD8DC",fontWeight:c.vencidoT>0?700:400,fontSize:12}}>{R(c.vencidoT)}</td>
                <td style={{padding:"9px 10px",textAlign:"right",color:c.v91>0?"#C62828":"#CFD8DC",fontWeight:c.v91>0?700:400,fontSize:12}}>{R(c.v91)}</td>
                <td style={{padding:"9px 10px",textAlign:"right",color:c.emit?"#6A1B9A":"#CFD8DC",fontWeight:c.emit?700:400,fontSize:12}}>{c.emit?`−${R(c.emit.total)}`:"-"}</td>
                <td style={{padding:"9px 10px"}}>
                  <span style={{background:c.st.bg,color:c.st.cor,fontWeight:700,fontSize:10,padding:"2px 8px",borderRadius:20,whiteSpace:"nowrap"}}>{c.st.l}</span>
                  {(()=>{
                    const allNfs=[...c.venc,...c.prox];
                    const tagged=allNfs.filter(n=>getNfStatus(execSel,c.g,n.nf));
                    if(!tagged.length) return null;
                    // show most critical status
                    const order=["protestado","aguardando","negociacao","conciliar","cobrado","pago"];
                    const top=order.find(s=>tagged.some(n=>getNfStatus(execSel,c.g,n.nf)===s));
                    if(!top) return null;
                    const sd=STATUS[top];
                    return <span style={{marginLeft:4,fontSize:10,fontWeight:700,color:sd.cor,background:sd.bg,padding:"2px 7px",borderRadius:20,border:`1px solid ${sd.cor}44`}}>{sd.e} {tagged.length}NF</span>;
                  })()}
                </td>
                <td style={{padding:"9px 10px"}}>
                  <button onClick={e=>{e.stopPropagation();setClientSel(c);setMsgT(c.emit?"conciliacao":"vencido");setSubTab("mensagens");}} style={{padding:"4px 10px",borderRadius:7,border:"none",cursor:"pointer",background:c.emit?"#EDE7F6":"#E3F2FD",color:c.emit?"#6A1B9A":"#0D47A1",fontWeight:700,fontSize:11}}>
                    {c.emit?"Conciliar":"Cobrar"}
                  </button>
                </td>
              </tr>
            ))}</tbody>
            <tfoot><tr style={{background:cor(execSel),color:"#fff",fontWeight:700}}>
              <td style={{padding:"10px"}} colSpan={3}>TOTAL ({execClients.length})</td>
              <td style={{padding:"10px",textAlign:"right"}}>{R(execClients.reduce((s,c)=>s+c.aV,0))}</td>
              <td style={{padding:"10px",textAlign:"right"}}>{R(execClients.reduce((s,c)=>s+c.vencidoT,0))}</td>
              <td style={{padding:"10px",textAlign:"right"}}>{R(execClients.reduce((s,c)=>s+c.v91,0))}</td>
              <td style={{padding:"10px",textAlign:"right"}}>−{R(Math.abs(execClients.reduce((s,c)=>s+(c.emit?.total||0),0)))}</td>
              <td colSpan={2}/>
            </tr></tfoot>
          </table>
        </div>
      </>}

      {/* ALERTAS */}
      {subTab==="alertas"&&<div style={{background:"#fff",borderRadius:12,padding:22,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
        <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>🔔 NFs Vencendo nos Próximos 7 Dias</div>
        {execAlerts.length===0?<div style={{color:"#90A4AE",textAlign:"center",padding:50}}>Sem vencimentos iminentes 🎉</div>
          :execAlerts.map((p,i)=>{
            const d=p.d; const badge=d===0?"HOJE":d===1?"AMANHÃ":`em ${d}d`; const bc=d===0?"#C62828":d<=2?"#EF6C00":"#1565C0";
            return(<div key={i} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 14px",borderRadius:10,background:"#F8FAFC",marginBottom:8,border:`1px solid ${p.hasEmit?"#E1BEE7":"#E3F2FD"}`}}>
              <div style={{textAlign:"center",background:d<=1?"#FFEBEE":"#E3F2FD",borderRadius:9,padding:"8px 12px",minWidth:52}}>
                <div style={{fontSize:17,fontWeight:800,color:bc,lineHeight:1}}>{p.dt.split("/")[0]}</div>
                <div style={{fontSize:9,color:bc}}>/{p.dt.split("/")[1]}</div>
              </div>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:13}}>{p.g} {p.hasEmit&&<span style={{fontSize:10,color:"#6A1B9A",fontWeight:700}}>💜 tem crédito EMIT</span>}</div>
                <div style={{fontSize:11,color:"#607D8B"}}>NF {p.nf} · {p.r}</div>
              </div>
              <div style={{fontWeight:800,color:"#1565C0",fontSize:14}}>{R(p.vlr)}</div>
              <span style={{background:d<=1?"#FFEBEE":"#E8F5E9",color:bc,fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20}}>{badge}</span>
              <button onClick={()=>{setClientSel(p.cRef);setMsgT(p.hasEmit?"conciliacao":"lancamento");setSubTab("mensagens");}} style={{padding:"7px 12px",borderRadius:8,border:"none",cursor:"pointer",fontWeight:700,fontSize:11,background:p.hasEmit?"#6A1B9A":"#1565C0",color:"#fff"}}>
                {p.hasEmit?"Conciliar":"Notificar"}
              </button>
            </div>);
          })
        }
      </div>}

      {/* CONCILIACAO */}
      {subTab==="conciliacao"&&<>
        <div style={{background:"#F3E5F5",border:"2px solid #CE93D8",borderRadius:12,padding:16,marginBottom:16,display:"flex",gap:12,alignItems:"center"}}>
          <span style={{fontSize:28}}>💜</span>
          <div>
            <div style={{fontWeight:800,fontSize:14,color:"#4A148C"}}>Créditos EMIT — {execSel.split(" ")[0]}</div>
            <div style={{fontSize:12,color:"#6A1B9A",marginTop:3}}>Clientes com créditos pendentes de conciliação no SAP. Não cobrar antes de conciliar.</div>
          </div>
          <div style={{marginLeft:"auto",textAlign:"right"}}>
            <div style={{fontSize:10,color:"#7B1FA2"}}>Total créditos</div>
            <div style={{fontWeight:800,fontSize:18,color:"#4A148C"}}>{R(execInfo.emit)}</div>
          </div>
        </div>
        {allClients[execSel].filter(c=>c.emit).map((c,i)=>(
          <div key={i} style={{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 6px rgba(0,0,0,.08)",marginBottom:12,border:"1px solid #E1BEE7"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12,flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontWeight:700,fontSize:14}}>{c.g}</div>
                <div style={{fontSize:12,color:"#607D8B"}}>{c.r} · {c.c}/{c.uf}</div>
                {(()=>{
                  const em  = getEmail(execSel,c);
                  const fr  = getFrom(execSel,c);
                  const deKey  = "de|"   + emailKey(execSel,c.g);
                  const paraKey = emailKey(execSel,c.g);
                  return (
                    <div style={{background:"#F8FAFC",borderRadius:8,padding:"7px 10px",marginTop:6,border:"1px solid #E3F2FD"}}>
                      <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:3}}>
                        <span style={{fontSize:10,color:"#90A4AE",fontWeight:600,minWidth:68}}>De:</span>
                        <span style={{fontSize:11,color:fr?"#2E7D32":"#B0BEC5"}}>{fr||"(não cadastrado)"}</span>
                        <button onClick={()=>{setClientSel(c);setEditingEmail(deKey);setEmailDraft(fr);setSubTab("mensagens");}}
                          style={{fontSize:10,background:"#E8F5E9",color:"#2E7D32",border:"none",borderRadius:5,padding:"1px 7px",cursor:"pointer"}}>✏️</button>
                        {fr&&<span style={{fontSize:9,background:"#E8F5E9",color:"#2E7D32",borderRadius:10,padding:"1px 5px",fontWeight:700}}>salvo</span>}
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <span style={{fontSize:10,color:"#90A4AE",fontWeight:600,minWidth:68}}>Para:</span>
                        <span style={{fontSize:11,color:em?"#1565C0":"#B0BEC5"}}>{em||"(não cadastrado)"}</span>
                        <button onClick={()=>{setClientSel(c);setEditingEmail(paraKey);setEmailDraft(em);setSubTab("mensagens");}}
                          style={{fontSize:10,background:"#E3F2FD",color:"#1565C0",border:"none",borderRadius:5,padding:"1px 7px",cursor:"pointer"}}>✏️</button>
                        {em&&<span style={{fontSize:9,background:"#E8F5E9",color:"#2E7D32",borderRadius:10,padding:"1px 5px",fontWeight:700}}>salvo</span>}
                      </div>
                    </div>
                  );
                })()}
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-start"}}>
                <div style={{textAlign:"right",background:"#FFF5F5",borderRadius:8,padding:"6px 12px"}}>
                  <div style={{fontSize:10,color:"#607D8B"}}>Vencido</div>
                  <div style={{fontWeight:700,color:"#C62828",fontSize:14}}>{R(c.vencidoT)}</div>
                </div>
                <div style={{textAlign:"right",background:"#F3E5F5",borderRadius:8,padding:"6px 12px"}}>
                  <div style={{fontSize:10,color:"#607D8B"}}>Crédito</div>
                  <div style={{fontWeight:700,color:"#6A1B9A",fontSize:14}}>−{R(c.emit.total)}</div>
                </div>
                <div style={{textAlign:"right",background:c.liquidoC<=0?"#E8F5E9":"#FFF3E0",borderRadius:8,padding:"6px 12px"}}>
                  <div style={{fontSize:10,color:"#607D8B"}}>Saldo</div>
                  <div style={{fontWeight:700,color:c.liquidoC<=0?"#2E7D32":"#E65100",fontSize:14}}>{c.liquidoC<=0?"Crédito":R(c.liquidoC)}</div>
                </div>
              </div>
            </div>
            {c.emit.alerta&&<div style={{background:"#FFF3E0",border:"1px solid #FFB74D",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:"#E65100",fontWeight:600}}>{c.emit.alerta}</div>}
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {c.emit.itens.map((it,j)=>(
                <div key={j} style={{background:"#FAF5FF",borderRadius:8,padding:"8px 12px",fontSize:11,border:"1px solid #E1BEE7",flex:1,minWidth:200}}>
                  <span style={{fontWeight:600}}>Ref. {it.ref}</span>
                  <span style={{marginLeft:6,background:"#E1BEE7",color:"#4A148C",padding:"1px 5px",borderRadius:8,fontSize:10}}>{it.tipo}</span>
                  <div style={{color:"#6A1B9A",fontWeight:700,marginTop:4}}>{R(it.vlr)}</div>
                  <div style={{color:"#90A4AE",fontSize:10,marginTop:2}}>{it.obs}</div>
                </div>
              ))}
            </div>
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <button onClick={()=>{setClientSel(c);setMsgT("conciliacao");setSubTab("mensagens");}} style={{flex:1,padding:"10px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:"#6A1B9A",color:"#fff"}}>💜 Ver Mensagem</button>
              {(()=>{const em=getEmail(execSel,c);const fr=getFrom(execSel,c);return(<button onClick={()=>openEmail(em,fr,c,"conciliacao")} title={em?`Enviar para ${em}`:"E-mail não cadastrado"} style={{flex:1,padding:"10px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:"#1565C0",color:"#fff"}}>✉️ {em?"Abrir E-mail":"Sem E-mail"}</button>);})()}
            </div>
          </div>
        ))}
      </>}

      {/* MENSAGENS */}
      {subTab==="mensagens"&&(
        <div style={{display:"grid",gridTemplateColumns:"260px 1fr",gap:16}}>
          <div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 1px 6px rgba(0,0,0,.08)",maxHeight:560,overflowY:"auto"}}>
            {allClients[execSel].map((c,i)=>(
              <div key={i} onClick={()=>setClientSel(c)} style={{padding:"10px 13px",cursor:"pointer",background:clientSel?.g===c.g?"#E3F2FD":"transparent",borderLeft:clientSel?.g===c.g?`4px solid ${cor(execSel)}`:"4px solid transparent",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #F5F5F5"}}>
                <div style={{minWidth:0,flex:1}}>
                  <div style={{fontWeight:600,fontSize:12,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{c.g} {c.emit&&"💜"}</div>
                  <div style={{fontSize:10,color:"#90A4AE"}}>{R(c.total)}</div>
                </div>
                <span style={{background:c.st.bg,color:c.st.cor,fontWeight:700,fontSize:9,padding:"2px 7px",borderRadius:20,whiteSpace:"nowrap",marginLeft:8}}>{c.st.l}</span>
              </div>
            ))}
          </div>
          {clientSel?(
            <div>
              <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 6px rgba(0,0,0,.08)",marginBottom:12}}>
                <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:14}}>
                  <div>
                    <div style={{fontSize:16,fontWeight:800}}>{clientSel.r}</div>
                    <div style={{fontSize:12,color:"#607D8B"}}>{clientSel.g} · {clientSel.c}/{clientSel.uf}</div>
                  {/* ── De/Para editável ── */}
                  {(()=>{
                    const paraKey = emailKey(execSel, clientSel.g);
                    const deKey   = "de|" + emailKey(execSel, clientSel.g);
                    const para    = getEmail(execSel, clientSel);
                    const de      = getFrom(execSel, clientSel);
                    const EditRow = ({label, value, editKey, placeholder, accentColor, onEdit}) => {
                        const isEd = editingEmail === editKey;
                        return (
                          <div style={{marginTop:5}}>
                            <span style={{fontSize:10,color:"#90A4AE",fontWeight:600,marginRight:4}}>{label}</span>
                            {isEd ? (
                              <div style={{display:"flex",gap:5,marginTop:3,alignItems:"center"}}>
                                <input autoFocus value={emailDraft} onChange={e=>setEmailDraft(e.target.value)}
                                  onKeyDown={e=>{if(e.key==="Enter")saveEmail(execSel,clientSel.g);if(e.key==="Escape")setEditingEmail(null);}}
                                  placeholder={placeholder}
                                  style={{flex:1,padding:"5px 9px",borderRadius:7,border:`1.5px solid ${accentColor}`,fontSize:12,outline:"none"}}/>
                                <button onClick={()=>saveEmail(execSel,clientSel.g)} style={{padding:"5px 11px",borderRadius:7,border:"none",cursor:"pointer",background:accentColor,color:"#fff",fontWeight:700,fontSize:12}}>Salvar</button>
                                <button onClick={()=>setEditingEmail(null)} style={{padding:"5px 9px",borderRadius:7,border:"none",cursor:"pointer",background:"#ECEFF1",color:"#607D8B",fontSize:12}}>✕</button>
                              </div>
                            ) : (
                              <span style={{fontSize:12,color:value?accentColor:"#B0BEC5"}}>
                                {value || `(não cadastrado)`}
                                <button onClick={onEdit} style={{marginLeft:6,fontSize:10,background:"#F5F5F5",color:"#607D8B",border:"none",borderRadius:5,padding:"1px 7px",cursor:"pointer"}}>✏️ Editar</button>
                                {!!value&&<span style={{marginLeft:5,fontSize:10,background:"#E8F5E9",color:"#2E7D32",borderRadius:10,padding:"1px 6px",fontWeight:700}}>salvo</span>}
                              </span>
                            )}
                          </div>
                        );
                      };
                    return (
                      <div style={{background:"#F8FAFC",borderRadius:9,padding:"10px 13px",marginTop:8,border:"1px solid #E3F2FD"}}>
                        <EditRow
                          label="De (remetente):"
                          value={de}
                          editKey={deKey}
                          placeholder="atendimento.cliente@brsupply.com.br"
                          accentColor="#2E7D32"
                          onEdit={()=>{setEditingEmail(deKey);setEmailDraft(de);}}
                        />
                        <EditRow
                          label="Para (destinatário):"
                          value={para}
                          editKey={paraKey}
                          placeholder="financeiro@cliente.com.br"
                          accentColor="#1565C0"
                          onEdit={()=>{setEditingEmail(paraKey);setEmailDraft(para);}}
                        />
                      </div>
                    );
                  })()}
                  {/* dummy to close old ternary */}
                  </div>
                  <span style={{background:clientSel.st.bg,color:clientSel.st.cor,fontWeight:700,fontSize:13,padding:"5px 14px",borderRadius:20,height:"fit-content"}}>{clientSel.st.l}</span>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8,marginBottom:12}}>
                  {[{l:"A Vencer",v:clientSel.aV,c:"#1565C0"},{l:"Até 30d",v:clientSel.v30,c:"#F9A825"},{l:"31–60d",v:clientSel.v60,c:"#EF6C00"},{l:"61–90d",v:clientSel.v90,c:"#D84315"},{l:"+91d",v:clientSel.v91,c:"#C62828"}].map(ag=>(
                    <div key={ag.l} style={{background:"#F8FAFC",borderRadius:7,padding:"8px 9px",borderTop:`3px solid ${ag.c}`}}>
                      <div style={{fontSize:9,color:"#607D8B"}}>{ag.l}</div>
                      <div style={{fontWeight:700,color:ag.c,fontSize:12}}>{R(ag.v)}</div>
                    </div>
                  ))}
                </div>
                {clientSel.emit&&<div style={{background:"#FAF5FF",border:"1.5px solid #CE93D8",borderRadius:9,padding:"10px 14px",marginBottom:10}}>
                  <div style={{fontWeight:700,color:"#4A148C",fontSize:12,marginBottom:6}}>💜 Crédito EMIT: −{R(clientSel.emit.total)}</div>
                  {clientSel.emit.alerta&&<div style={{fontSize:11,color:"#C62828",fontWeight:600,marginBottom:6}}>{clientSel.emit.alerta}</div>}
                  {clientSel.emit.itens.map((it,i)=><div key={i} style={{fontSize:11,color:"#607D8B",marginTop:4}}>Ref. {it.ref} ({it.tipo}) — {R(it.vlr)}</div>)}
                </div>}
                {clientSel.venc.length>0&&<div style={{background:"#FFF5F5",borderRadius:8,padding:11,marginBottom:8}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#C62828",marginBottom:6}}>NFs Vencidas:</div>
                  {clientSel.venc.map((v,i)=>{
                    const st=getNfStatus(execSel,clientSel.g,v.nf);
                    const stDef=st?STATUS[st]:null;
                    return(
                      <div key={i} style={{marginBottom:6,padding:"6px 8px",borderRadius:7,background:stDef?stDef.bg:"transparent",border:stDef?`1px solid ${stDef.cor}33`:"1px solid transparent"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:6}}>
                          <span style={{fontSize:11,fontWeight:600}}>NF {v.nf} · {v.dt}</span>
                          <span style={{color:"#C62828",fontWeight:700,fontSize:11,flexShrink:0}}>{R(v.vlr)} <span style={{color:"#90A4AE",fontWeight:400}}>({v.d}d)</span></span>
                        </div>
                        <div style={{display:"flex",gap:4,marginTop:5,flexWrap:"wrap",alignItems:"center"}}>
                          {stDef&&<span style={{fontSize:10,fontWeight:700,color:stDef.cor,background:stDef.bg,padding:"2px 8px",borderRadius:20,border:`1px solid ${stDef.cor}55`}}>{stDef.e} {stDef.l}</span>}
                          <select value={st||""} onChange={e=>setNfStatusVal(execSel,clientSel.g,v.nf,e.target.value||null)}
                            style={{fontSize:10,padding:"2px 6px",borderRadius:6,border:"1px solid #CFD8DC",background:"#fff",cursor:"pointer",color:"#607D8B"}}>
                            <option value="">＋ Definir status</option>
                            {STATUS_LIST.map(([k,s])=><option key={k} value={k}>{s.e} {s.l}</option>)}
                          </select>
                        </div>
                      </div>
                    );
                  })}
                </div>}
                {clientSel.prox.length>0&&<div style={{background:"#E8F5E9",borderRadius:8,padding:11}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#2E7D32",marginBottom:6}}>Próximos Vencimentos:</div>
                  {clientSel.prox.map((p,i)=>{
                    const st=getNfStatus(execSel,clientSel.g,p.nf);
                    const stDef=st?STATUS[st]:null;
                    return(
                      <div key={i} style={{marginBottom:6,padding:"6px 8px",borderRadius:7,background:stDef?stDef.bg:"transparent",border:stDef?`1px solid ${stDef.cor}33`:"1px solid transparent"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:6}}>
                          <span style={{fontSize:11,fontWeight:600}}>NF {p.nf} · {p.dt}</span>
                          <span style={{color:"#2E7D32",fontWeight:700,fontSize:11,flexShrink:0}}>{R(p.vlr)}</span>
                        </div>
                        <div style={{display:"flex",gap:4,marginTop:5,flexWrap:"wrap",alignItems:"center"}}>
                          {stDef&&<span style={{fontSize:10,fontWeight:700,color:stDef.cor,background:stDef.bg,padding:"2px 8px",borderRadius:20,border:`1px solid ${stDef.cor}55`}}>{stDef.e} {stDef.l}</span>}
                          <select value={st||""} onChange={e=>setNfStatusVal(execSel,clientSel.g,p.nf,e.target.value||null)}
                            style={{fontSize:10,padding:"2px 6px",borderRadius:6,border:"1px solid #CFD8DC",background:"#fff",cursor:"pointer",color:"#607D8B"}}>
                            <option value="">＋ Definir status</option>
                            {STATUS_LIST.map(([k,s])=><option key={k} value={k}>{s.e} {s.l}</option>)}
                          </select>
                        </div>
                      </div>
                    );
                  })}
                </div>}
              </div>
              <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
                <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>Gerar Mensagem</div>
                <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>
                  {[["preventivo","✉️ Preventiva"],["lancamento","📄 Lançamento"],["vencido","⚠️ Cobrança"],["conciliacao","💜 Conciliação"]].map(([k,l])=>(
                    <button key={k} onClick={()=>setMsgT(k)} style={{padding:"7px 12px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:600,fontSize:12,background:msgT===k?cor(execSel):"#E3F2FD",color:msgT===k?"#fff":"#1565C0"}}>{l}</button>
                  ))}
                </div>
                <textarea readOnly value={MSGS[msgT](clientSel)} style={{width:"100%",minHeight:190,padding:12,borderRadius:9,border:"1.5px solid #CFD8DC",fontSize:12,resize:"vertical",fontFamily:"inherit",color:"#37474F",background:"#F8FAFC",boxSizing:"border-box",lineHeight:1.6}}/>
                <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
                  <button onClick={()=>copyMsg(clientSel)} style={{flex:1,minWidth:120,padding:"11px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:copied?"#2E7D32":cor(execSel),color:"#fff",transition:"background .3s"}}>
                    {copied?"✅ Copiado!":"📋 Copiar"}
                  </button>
                  {(()=>{
                    const em = getEmail(execSel, clientSel);
                    return (
                      <button
                        onClick={()=>{ const fr=getFrom(execSel,clientSel); openEmail(em, fr, clientSel, msgT); setLog(p=>[{id:Date.now(),exec:execSel,g:clientSel.g,acao:`E-mail aberto (${msgT}) → ${em||"sem e-mail"}`,ts:new Date().toLocaleString("pt-BR")},...p.slice(0,99)]); }}
                        title={em ? `Enviar para ${em}` : "E-mail não cadastrado — será aberto em branco"}
                        style={{flex:1,minWidth:120,padding:"11px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:"#1565C0",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                        ✉️ {em ? `Abrir E-mail` : "Sem e-mail"}
                      </button>
                    );
                  })()}
                  <button onClick={()=>setLog(p=>[{id:Date.now(),exec:execSel,g:clientSel.g,acao:"Contato registrado manualmente",ts:new Date().toLocaleString("pt-BR")},...p.slice(0,99)])} style={{padding:"11px 14px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:"#E8F5E9",color:"#2E7D32"}}>✓ Registrar</button>
                </div>
                {!getEmail(execSel,clientSel)&&<div style={{marginTop:8,fontSize:11,color:"#EF6C00",background:"#FFF3E0",borderRadius:7,padding:"6px 10px"}}>⚠️ E-mail não cadastrado. Clique em <strong>✏️ Editar</strong> na ficha do cliente para adicionar.</div>}
              </div>
            </div>
          ):(
            <div style={{background:"#fff",borderRadius:12,padding:50,textAlign:"center",boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{fontSize:40,marginBottom:12}}>👈</div>
              <div style={{fontWeight:600,color:"#607D8B"}}>Selecione um cliente</div>
            </div>
          )}
        </div>
      )}

      {/* STATUS */}
      {subTab==="status"&&(()=>{
        // Collect all NFs that have a status for this exec
        const rows=[];
        (allClients[execSel]||[]).forEach(c=>{
          [...c.venc,...c.prox].forEach(nf=>{
            const st=getNfStatus(execSel,c.g,nf.nf);
            if(st) rows.push({cliente:c.g,razao:c.r,nf:nf.nf,dt:nf.dt,vlr:nf.vlr,dias:nf.d||null,st,c});
          });
        });
        // Group by status
        const byStatus={};
        STATUS_LIST.forEach(([k])=>{ byStatus[k]=rows.filter(r=>r.st===k); });
        const total=rows.length;
        return(
          <div>
            {/* Summary cards */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:18}}>
              {STATUS_LIST.map(([k,s])=>(
                <div key={k} style={{background:"#fff",borderRadius:10,padding:"12px 14px",boxShadow:"0 1px 6px rgba(0,0,0,.07)",borderTop:`4px solid ${s.cor}`,textAlign:"center"}}>
                  <div style={{fontSize:20,marginBottom:4}}>{s.e}</div>
                  <div style={{fontWeight:800,fontSize:22,color:s.cor}}>{byStatus[k].length}</div>
                  <div style={{fontSize:11,color:"#607D8B",marginTop:2}}>{s.l}</div>
                </div>
              ))}
            </div>

            {total===0?(
              <div style={{background:"#fff",borderRadius:12,padding:60,textAlign:"center",boxShadow:"0 1px 6px rgba(0,0,0,.07)"}}>
                <div style={{fontSize:40,marginBottom:12}}>🏷️</div>
                <div style={{fontWeight:600,color:"#607D8B"}}>Nenhuma NF com status atribuído ainda.</div>
                <div style={{fontSize:12,color:"#90A4AE",marginTop:6}}>Acesse a aba Mensagens e defina o status de cada NF.</div>
              </div>
            ):(
              <div>
                {STATUS_LIST.map(([k,s])=> byStatus[k].length===0?null:(
                  <div key={k} style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 6px rgba(0,0,0,.07)",marginBottom:14}}>
                    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
                      <span style={{fontSize:20}}>{s.e}</span>
                      <div style={{fontWeight:700,fontSize:15,color:s.cor}}>{s.l}</div>
                      <span style={{background:s.bg,color:s.cor,fontWeight:700,fontSize:12,padding:"2px 10px",borderRadius:20,border:`1px solid ${s.cor}44`}}>{byStatus[k].length} NF{byStatus[k].length>1?"s":""}</span>
                      <span style={{marginLeft:"auto",fontSize:12,color:"#607D8B",fontWeight:600}}>
                        Total: {byStatus[k].reduce((a,r)=>a+r.vlr,0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"})}
                      </span>
                    </div>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:s.bg}}>
                        {["Cliente","NF","Vencimento","Valor","Atraso","Ação"].map(h=>(
                          <th key={h} style={{padding:"8px 10px",textAlign:h==="Valor"||h==="Atraso"?"right":"left",color:s.cor,fontWeight:700,fontSize:11}}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {byStatus[k].map((r,i)=>(
                          <tr key={i} style={{borderBottom:"1px solid #F5F5F5"}} onMouseEnter={e=>e.currentTarget.style.background="#FAFAFA"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                            <td style={{padding:"9px 10px",fontWeight:600}}>{r.cliente}<div style={{fontSize:10,color:"#90A4AE",fontWeight:400}}>{r.razao.slice(0,35)}</div></td>
                            <td style={{padding:"9px 10px",fontFamily:"monospace"}}>{r.nf}</td>
                            <td style={{padding:"9px 10px",color:"#607D8B"}}>{r.dt}</td>
                            <td style={{padding:"9px 10px",textAlign:"right",fontWeight:700}}>{R(r.vlr)}</td>
                            <td style={{padding:"9px 10px",textAlign:"right",color:r.dias?"#C62828":"#90A4AE"}}>{r.dias?`${r.dias}d`:"—"}</td>
                            <td style={{padding:"9px 10px"}}>
                              <div style={{display:"flex",gap:5,alignItems:"center"}}>
                                <select value={r.st} onChange={e=>setNfStatusVal(execSel,r.cliente,r.nf,e.target.value||null)}
                                  style={{fontSize:10,padding:"3px 6px",borderRadius:6,border:`1px solid ${s.cor}55`,background:s.bg,cursor:"pointer",color:s.cor,fontWeight:600}}>
                                  <option value="">— remover —</option>
                                  {STATUS_LIST.map(([sk,ss])=><option key={sk} value={sk}>{ss.e} {ss.l}</option>)}
                                </select>
                                <button onClick={()=>{setClientSel(r.c);setSubTab("mensagens");}} style={{fontSize:10,padding:"3px 8px",borderRadius:6,border:"none",cursor:"pointer",background:"#E3F2FD",color:"#1565C0",fontWeight:600}}>Ver</button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      {/* HISTÓRICO */}
      {subTab==="historico"&&(
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:14}}>Histórico de Ações</div>
          {log.filter(l=>l.exec===execSel).length===0
            ?<div style={{color:"#90A4AE",textAlign:"center",padding:50}}><div style={{fontSize:36,marginBottom:10}}>📋</div>Nenhuma ação registrada.</div>
            :log.filter(l=>l.exec===execSel).map(a=>(
              <div key={a.id} style={{display:"flex",gap:12,padding:"10px 0",borderBottom:"1px solid #ECEFF1",alignItems:"center"}}>
                <span style={{fontSize:18}}>📌</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:600,fontSize:13}}>{a.g}</div>
                  <div style={{fontSize:11,color:"#607D8B"}}>{a.acao}</div>
                </div>
                <div style={{fontSize:11,color:"#90A4AE"}}>{a.ts}</div>
              </div>
            ))
          }
        </div>
      )}

      </div>
    </div>
  );
}
