var Docx = require("docxCreater");
const path = require('path');

var str = `黃應怕個斯少油統管即不
我太識統，假那發，五進因聲、野持會平的歌法她西同經對形，準之前如五多筆她？開軍車人長？因座利車德他黑然了天統部找物收中標拉術中的面的文了，現腦只最停之花廠中，天交電前往山片友戰人象了數區生有人香為會期時臺工奇友話造辦象……題為間大情回期，謝最益模最美雨都？有買驚體分：身去得自親建一看天。是歡亮產市後的，我活保先經性，線答他叫爸西真倒重人馬光機，即起去縣展往，多先價兩真標到雄有，雜這克靈就，這己照血效斯說是就得邊。

老行然散，決依面北廣，必算道市會到來、不父美望場其情所之少生臺個作媽，也基否學的委真國線道什一女民王上！的和一外……升只一卻資利會天到石理生好問，優顯前。

樣活步經極根對是！地世真新大會黨，數公都前體王今兒人知微意那明食衣才就得看寫、教改除港自道中光雙實司那散感應則銀西電容產只、治傳只早多但縣水廣極樣後不！術時西……眼期。

再土的藥費人銷利統的對，死氣會看有媽後基科此背先找存步這古我一千數統？人地著上停果不起，臺在白員造對體但一照……觀只留音了心等首等起古燈校要與設，我院優如看機、經以怎以我老舉林得沒家河絕化門、動賣育見音麼，入裡教前場史黨來？

程縣其。會的開員對功、用時的國大包的高關強古教直反看二錯里爸我時海陽！轉外的了！究大表：長管條速生；興現股能，主這術大出青望理種品請人意畫，法他流打運少書，大傷理少。程調如把產進房懷分，來兩呢年只果，會生加質特立媽興……麼本興出形好素非課留寫於化記況血友企我依則行難我且情創身後平大少……那是制……鄉可神。識都嗎，想門原。

功至學兒確說年清局教人現管使修直我：農女知保科熱國在樂論止春得招。紀分在收就內父，投不上的到舉人雖立們手實，市從點有？分必市元作著試能為已怎，心天了點富眼一，性人學寫家式自。點工明驚，五成了全愛轉我本政覺作覺特其教傳師活上竟氣們所禮……引方多入優下花北一正，首取的視了一！產人個老山年河本打老知輕家方打系年畫賽能山驚課吃而、別王招？木看信紀夫檢又定色主動不嚴：不何。

這手地要時人優足怕軍，書數友懷倒係我、裡在學就因自放，些風爾來影得起傷劇布時學大同所是聯太但學年行。

的好幾海們顧，心行有常平停單變兒想學團器走，得了經……樂媽產使像需太、本在河爸車我自交半愛……據麼三加品不！門個以身遠環覺入和在大時雖斷、時大童產式西人定卻感影把有速面無少經答本去關方力會？國有甚半狀還？意著候的立英公程加，河不化唱不場處計同西是法起不外有吸歷吃人布夠分爾聲於建量多，身發一字專流雖一怎專政光看表程王司智問結可了成來；處地此考配死太名人果子他？為獨接月東的直值利，品業會心年直。山利質聞用提認她重有自辦紅土然前……心紀關不活作人在五學因上自事了，整開交是望來技代重由科發率機一此因奇長衣選圖以野我的再便些意當？進動畫的進解來種痛，當起法對安易成輕就書心政有市施能。事南年而成車上文談因收其。`


// str = str.repeat(20);





/*
 var d = new Docx();
 d.createParagraph('hello word').addText("hello world 2").addLink("hello world3","http://www.baidu.com")
 d.createParagraph(str).addImage("C:\\Users\\Administrator\\Desktop\\测试图片类型\\111xiaotupian.jpg",55,55,"aboveText").addText(str)
 d.createParagraph().addImage("C:\\Users\\Administrator\\Desktop\\测试图片类型\\111.png",55,55)
 d.createParagraph().addImage("C:\\Users\\Administrator\\Desktop\\测试图片类型\\111.gif",55,55)
 d.createParagraph().addImage("C:\\Users\\Administrator\\Desktop\\测试图片类型\\222.gif",55,55)
 d.render()*/

var d = new Docx({
    'filePath':path.join(__dirname,'testMathXML.docx')
});

d.addParagraph(str).setPage({type:'A3',orient:true,cols:2}).addText(`
            <m:oMathPara>
                <m:oMath>
                    <m:f>
                        <m:fPr>
                            <m:ctrlPr>
                                <w:rPr>
                                    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                </w:rPr>
                            </m:ctrlPr>
                        </m:fPr>
                        <m:num>
                            <m:r>
                                <m:rPr>
                                    <m:sty m:val="p"/>
                                </m:rPr>
                                <w:rPr>
                                    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                </w:rPr>
                                <m:t>-b±</m:t>
                            </m:r>
                            <m:rad>
                                <m:radPr>
                                    <m:degHide m:val="1"/>
                                    <m:ctrlPr>
                                        <w:rPr>
                                            <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                        </w:rPr>
                                    </m:ctrlPr>
                                </m:radPr>
                                <m:deg/>
                                <m:e>
                                    <m:sSup>
                                        <m:sSupPr>
                                            <m:ctrlPr>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                                </w:rPr>
                                            </m:ctrlPr>
                                        </m:sSupPr>
                                        <m:e>
                                            <m:r>
                                                <m:rPr>
                                                    <m:sty m:val="p"/>
                                                </m:rPr>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                                </w:rPr>
                                                <m:t>b</m:t>
                                            </m:r>
                                        </m:e>
                                        <m:sup>
                                            <m:r>
                                                <m:rPr>
                                                    <m:sty m:val="p"/>
                                                </m:rPr>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                                </w:rPr>
                                                <m:t>2</m:t>
                                            </m:r>
                                        </m:sup>
                                    </m:sSup>
                                    <m:r>
                                        <m:rPr>
                                            <m:sty m:val="p"/>
                                        </m:rPr>
                                        <w:rPr>
                                            <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                        </w:rPr>
                                        <m:t>-4ac</m:t>
                                    </m:r>
                                </m:e>
                            </m:rad>
                        </m:num>
                        <m:den>
                            <m:r>
                                <m:rPr>
                                    <m:sty m:val="p"/>
                                </m:rPr>
                                <w:rPr>
                                    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                                </w:rPr>
                                <m:t>2a</m:t>
                            </m:r>
                        </m:den>
                    </m:f>
                </m:oMath>
            </m:oMathPara>`);



d.render()