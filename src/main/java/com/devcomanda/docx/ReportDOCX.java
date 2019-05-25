package com.devcomanda.docx;

import com.devcomanda.recipe.Recipe;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import java.io.*;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.Map;

public class ReportDOCX {

    private String filePath;

    public ReportDOCX(String filePath) {
        this.filePath = filePath;
    }

    public void create() throws IOException, InvalidFormatException {
        Recipe recipe = generateRecipe();

        XWPFDocument document = new XWPFDocument();

        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun runTitle = title.createRun();
        runTitle.setText(recipe.getName());
        runTitle.setFontSize(18);
        runTitle.setBold(true);

        printImage(document, recipe.getUrlToImage());

        printIngredients(document, recipe.getIngredients());

        printItems(document, recipe.getItems());

        document.write(new FileOutputStream(new File(filePath)));
        document.close();
    }

    private void printImage(XWPFDocument document, String urlImage) throws IOException, InvalidFormatException {
        FileInputStream is = new FileInputStream(urlImage);

        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = paragraph.createRun();
        run.addBreak();
        run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, "chizkejk.jpeg", Units.toEMU(300), Units.toEMU(225));

        is.close();
    }

    private void printIngredients(XWPFDocument document, Iterator<Map.Entry<String, String>> ingredients) {
        CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
        cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));

        CTLvl cTLvl = cTAbstractNum.addNewLvl();
        cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
        cTLvl.addNewLvlText().setVal("•");

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);

        XWPFNumbering numbering = document.createNumbering();

        BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);

        BigInteger numID = numbering.addNum(abstractNumID);

        while (ingredients.hasNext()) {
            Map.Entry<String, String> ingredient = ingredients.next();

            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setNumID(numID);
            paragraph.setSpacingAfter(0);

            XWPFRun name = paragraph.createRun();
            name.setBold(true);
            name.setFontSize(14);
            name.setColor("96d228");
            name.setText(ingredient.getKey());

            paragraph.createRun().setText(" - ");

            XWPFRun amount = paragraph.createRun();
            amount.setItalic(true);
            amount.setFontSize(14);
            amount.setText(ingredient.getValue());
        }
    }

    private void printItems(XWPFDocument document, Iterator<String> items) {

        while (items.hasNext()) {
            String item = items.next();

            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.BOTH);
            paragraph.setSpacingBefore(10);

            XWPFRun run = paragraph.createRun();
            run.setFontSize(14);
            run.setText(item);
        }
    }

    private Recipe generateRecipe() {
        Recipe recipe = new Recipe("Чизкейк");

        recipe.addIngredient("Печенье", "400 г");
        recipe.addIngredient("Масло сливочное", "150 г");
        recipe.addIngredient("Творог", "800 г");
        recipe.addIngredient("Сливки", "120 г");
        recipe.addIngredient("Сахар", "200 г");
        recipe.addIngredient("Яйцо куриное", "3 шт");
        recipe.addIngredient("Ванильный сахар", "1 упак.");
        recipe.addIngredient("Ванилин", "по вкусу");
        recipe.addIngredient("Малина", "400 г");
        recipe.addIngredient("Сметана", "250 г");

        recipe.addItem("Данное количество ингредиентов рассчитано на разъемную форму диаметром 23 см, получается высокий чизкейк с бортиками.");
        recipe.addItem("Измельчаем печенье в крошку.");
        recipe.addItem("Растапливаем сливочное масло в микроволновой печи. У меня этот процесс занял 45 секунд (на мощности 800 Вт).");
        recipe.addItem("Пока масло остывает, застилаем дно формы пергаментной бумагой. Съемную часть формы я не застилаю пергаментом, бортики никогда не пригорают и не прилипают, я на всякий случай провожу ножом между формой и основой, а потом легко снимаю форму. Советую не обрезать края пергаментой бумаги, тогда будет легче переместить готовый чизкейк на блюдо для подачи.");
        recipe.addItem("Вливаем растопленное сливочное масло в емкость с печеньем и тщательно перемешиваем ингредиенты.");
        recipe.addItem("Утрамбовываем массу в разъемной форме. Толщина дна чизкейка – чуть более 10 мм, бортиков – примерно 5-7 мм.");
        recipe.addItem("Включаем духовку на 170 градусов.");
        recipe.addItem("Отправляем основу для чизкейка в холодильник и переходим к самому вкусному – к начинке.");
        recipe.addItem("Протираем творог через сито или «перетираем» блендером до изменения его текстуры в однородную, почти кремовую. С использованием блендера у меня на это уходит около 4-х минут.");
        recipe.addItem("Начинаем превращение творога в некое подобие крем-чиза – добавляем сливки и снова «перетираем» блендером до приобретения массой глянцевой и нежной текстуры.");
        recipe.addItem("Добавляем в полученную массу 3 яйца, 200 г сахара и упаковку ванильного сахара (у меня в упаковке – 10 г). Снова «перетираем» все ингредиенты блендером, стараясь избегать чрезмерного насыщения творожной массы воздухом – перемещаем блендер внутри массы, как можно реже выводя его на поверхность.");
        recipe.addItem("Вынимаем из холодильника основу для чизкейка и равномерно наполняем ее творожной массой. После проведения вышеуказанных манипуляций я знаю, что нужно стукнуть формой по столу, чтоб избавить чизкейк от лишнего воздуха. Не понимаю механизм работы сего действия и не знаю, действительно ли это помогает, но каждый раз повторяю этот прием.");
        recipe.addItem("Отправляем чизкейк в разогретую до 170 градусов духовку и оставляем его в одиночестве на 50 минут. Чизкейк вполне предсказуемо ведет себя на протяжение 50-ти минут – не растекается, не поднимается, не подгорает.");
        recipe.addItem("Взбиваем сметану с 2-мя ст. л. сахара и ванилином до однородности.");
        recipe.addItem("По истечении 50-ти минут, вынимаем чизкейк из духовки. Вы заметите, что масса существенно уплотнилась.");
        recipe.addItem("Увеличиваем температуру духовки до 200 градусов. Выливаем сметанную массу поверх творожной, если необходимо, слегка «пригладьте» ее ложкой. Отправляем чизкейк в разогретую до 200 градусов духовку на 7-мь минут.");
        recipe.addItem("По истечении 7-ми минут достаем чизкейк из духовки, украшаем ягодами и отправляем в холодильник на всю ночь.");
        recipe.addItem("Угощайтесь!");

        recipe.setUrlToImage("E:\\ApachePOITest\\chizkejk.jpeg");

        return recipe;
    }
}
