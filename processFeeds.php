<?php
include $_SERVER["DOCUMENT_ROOT"] . '/local/vendor/autoload.php';
require_once ($_SERVER['DOCUMENT_ROOT'] . "/bitrix/modules/main/include/prolog_before.php");
CModule::IncludeModule('iblock');
\Bitrix\Main\Loader::includeModule('catalog');
use PhpOffice;
$inputFileName = $_SERVER['DOCUMENT_ROOT'] . "/local/feeds/" . $_GET['PATH'];

class processFeeds
{

    /* Настройки */

    protected $IBLOCK_ID = 4; //Инфоблок модификаций
    protected $createProps = true; // Создавать новые свойства
    protected $reservedNames = ['Категория', 'Название', 'Цена', 'Ед', 'Фотография', 'Вес', 'Описание', 'Доп. Фото'];

    /* Конец настроек */

    protected $itemsSheet = [];

    protected $propertyCollection = [];

    protected $measures = [];

    protected $sectId;

    protected $sellerName;

    //Парсинг всех столбцов и упаковка их в ассоциативный массив
    private function prepareArrayItems($path, $limit, $start)
    {

        $spreadsheet = PhpOffice\PhpSpreadsheet\IOFactory::load($path);
        $worksheet = $spreadsheet->getActiveSheet();

        $header_values = $sheet_array = [];

        $currentChecked = 0;

        $limit = $limit + $start;

        foreach ($worksheet->getRowIterator() AS $k => $row)
        {

            if ($k === 1)
            {
                foreach ($row->getCellIterator() as $cell)
                {
                    $header_values[] = $cell->getValue();
                }
                continue;
            }

            if ($currentChecked < $start)
            {
                $currentChecked++;
                continue;
            }

            $content_values = [];
            foreach ($row->getCellIterator() as $cell)
            {
                $content_values[] = $cell->getValue();
            }
            $sheet_array[] = array_combine($header_values, $content_values);

            //Лимит на парсинг
            $currentChecked++;
            if ($currentChecked >= $limit)
            {
                break;
            }

        }

        //Получение единиц измерения
        $mList = CCatalogMeasure::getList();

        while ($arResult = $mList->GetNext())
        {
            $this->measures[$arResult['SYMBOL_RUS']] = $arResult['ID'];
        }

        return $sheet_array;
    }

    //Обработка солбцов
    private function processItems()
    {

        foreach ($this->itemsSheet as $item)
        {

            //Если пустая категория
            if (empty($item['Категория'])) continue;

            //Если не заполнена базовая информация
            if (!$this->checkMainInfo($item)) continue;

            //Получение свойств товара
            $this->getProperties($item['Категория']);

            //Создание товара
            $this->addProduct($item);

        }

    }

    private function addProduct($item)
    {

        //Поиск существующего товара
        $productID = $this->searchModification($item['Название']);

        //echo $item['Категория'];
        //Если товар уже существует
        if ($productID > 0) $this->createOffer($item, $productID); // Добавляем торговое предложение
        else $this->createModification($item);

        echo 'Product id ' . $productID;

    }

    //Поиск модификаций
    private function searchModification($name)
    {
        $arSelect = Array(
            "ID"
        );
        $arFilter = Array(
            "IBLOCK_ID" => 4,
            "CODE" => $this->translit($name)
        );
        $res = CIBlockElement::GetList(Array() , $arFilter, false, Array() , $arSelect);
        if ($ob = $res->GetNextElement())
        {
            $arFields = $ob->GetFields();
            return $arFields['ID'];
        }
        return 0;
    }

    //Создание товара
    private function createModification($item)
    {
        $props = [];

        $props = $this->preparePropertiesArray($item);

        echo '<pre> Пропсы', print_r($props) , '</pre>';

        $photo = CFile::MakeFileArray($item['Фотография']);

        $addPhoto = CFile::MakeFileArray($item['Доп. Фото']);

        $props['MORE_PHOTO'] = $addPhoto;

        $el = new CIBlockElement;

        $arLoadProductArray = Array(
            "IBLOCK_SECTION_ID" => $this->sectId,
            "IBLOCK_ID" => $this->IBLOCK_ID,
            "PROPERTY_VALUES" => $props,
            "NAME" => $item['Название'],
            "CODE" => $this->translit($item['Название']) ,
            "ACTIVE" => "Y", // активен
            "DETAIL_PICTURE" => $photo,
            "DETAIL_TEXT" => $item['Описание']
        );
        $productId = $el->Add($arLoadProductArray);
        if ($productId)
        {
            echo $item['Название'] . ' создана модификация ID: ' . $productId;
            $this->createOffer($item, $productId);
        }
        else
        {
            echo 'Ошибка при добавлении товара ' . $el->LAST_ERROR;
        }
    }

    //Создание торгового предложения
    private function createOffer($item, $productID)
    {
        $props = array(
            'CML2_LINK' => $productID,
            'ATT_seller_name' => $this->sellerName,
            'ATT_price' => round($item['Цена'])
        );
        $el = new CIBlockElement;

        $arLoadProductArray = Array(
            "IBLOCK_SECTION_ID" => $this->getOffersSect($item['Категория']) ,
            "IBLOCK_ID" => 5,
            "CREATED_BY" => 31,
            "PROPERTY_VALUES" => $props,
            "NAME" => $item['Название'],
            "ACTIVE" => "Y", // активен

        );
        if ($offerID = $el->Add($arLoadProductArray))
        {
            $fields = array(
                'ID' => $offerID,
                'QUANTITY_TRACE' => \Bitrix\Catalog\ProductTable::STATUS_DEFAULT,
                'CAN_BUY_ZERO' => \Bitrix\Catalog\ProductTable::STATUS_DEFAULT,
                'WEIGHT' => $item['Вес'],
                'MEASURE' => $this->measures[$item['Ед']],
                'WIDTH' => $item['Ширина'],
                'LENGTH' => $item['Длина'],
                'HEIGHT' => $item['Толщина']
            );
            // создание товара
            $result = CCatalogProduct::Add($fields);
            CPrice::SetBasePrice($offerID, round($item['Цена']) , "RUB");
        }
        else echo "Error offer: " . $el->LAST_ERROR;

    }

    //Получить ID категории ТП
    private function getOffersSect($name)
    {
        $rsSect = CIBlockSection::GetList(array() , array(
            'NAME' => $name,
            'IBLOCK_ID' => 5
        ));

        if ($arSect = $rsSect->GetNext())
        {
            return $arSect['ID'];
        }
    }

    //Перевод текста в транслит
    private function translit($s)
    {
        $s = (string)$s; // преобразуем в строковое значение
        $s = trim($s); // убираем пробелы в начале и конце строки
        $s = function_exists('mb_strtolower') ? mb_strtolower($s) : strtolower($s); // переводим строку в нижний регистр (иногда надо задать локаль)
        $s = strtr($s, array(
            'а' => 'a',
            'б' => 'b',
            'в' => 'v',
            'г' => 'g',
            'д' => 'd',
            'е' => 'e',
            'ё' => 'e',
            'ж' => 'j',
            'з' => 'z',
            'и' => 'i',
            'й' => 'y',
            'к' => 'k',
            'л' => 'l',
            'м' => 'm',
            'н' => 'n',
            'о' => 'o',
            'п' => 'p',
            'р' => 'r',
            'с' => 's',
            'т' => 't',
            'у' => 'u',
            'ф' => 'f',
            'х' => 'h',
            'ц' => 'c',
            'ч' => 'ch',
            'ш' => 'sh',
            'щ' => 'shch',
            'ы' => 'y',
            'э' => 'e',
            'ю' => 'yu',
            'я' => 'ya',
            'ъ' => '',
            'ь' => '',
            ' ' => '_',
            '(' => '_',
            ')' => '_',
            '/' => '_',
            ',' => '_'
        ));
        return $s; // возвращаем результат

    }

    //Подготовка массива свойств
    private function preparePropertiesArray($item)
    {
        $props = [];

        foreach ($item as $propName => $propValue)
        {
            if (empty($propValue)) continue;
            if (!empty($this->propertyCollection[$item['Категория']][$propName]))
            {
                $code = $this->propertyCollection[$item['Категория']][$propName]['CODE'];
                if ($this->propertyCollection[$item['Категория']][$propName]['TYPE'] === 'L')
                {
                    if (!empty($this->propertyCollection[$item['Категория']][$propName]['VALUES'][$propValue]))
                    {
                        $props[$code] = $this->propertyCollection[$item['Категория']][$propName]['VALUES'][$propValue]['CODE'];
                    }
                    else
                    {
                        $valueCode = $this->findOrCreateProp($propName, $propValue, $item['Категория'], true);
                        $props[$code] = $valueCode;
                    }
                }
                else
                {
                    $props[$code] = $propValue;
                }
            }
            else if (empty($this->propertyCollection[$item['Категория']][$propName]) && !in_array($propName, $this->reservedNames) && $this->createProps)
            {
                echo 'Создаю новое свойство ' . $propName;
                //Поиск свойства и его привязка
                $res = $this->findOrCreateProp($propName, $propValue, $item['Категория']);
                $props[$res['CODE']] = $res['VALUE'];
            }
        }
        return $props;
    }

    private function findOrCreateProp($propName, $propValue, $categoryName, $findOnlyValue = false)
    {
        $properties = CIBlockProperty::GetList(Array(
            "sort" => "asc",
            "name" => "asc"
        ) , Array(
            "ACTIVE" => "Y",
            "NAME" => $propName,
            "IBLOCK_ID" => $this->IBLOCK_ID
        ));
        if ($prop = $properties->GetNext())
        {

            //Если задача - найти только значение
            if ($findOnlyValue)
            {
                $valueCode = $this->findOrCreateValue($prop['ID'], $propValue);
                $this->propertyCollection[$categoryName][$propName]['VALUES'] = $this->addToPropertyCollectionList($prop['CODE']);
                return $valueCode;

            }

            $valueID = $this->findOrCreateValue($prop['ID'], $propValue);
            CIBlockSectionPropertyLink::Add($this->sectId, $prop['ID'], $arLink = array(
                'IBLOCK_ID' => $this->IBLOCK_ID,
                'SMART_FILTER' => 'Y'
            )); // Привязываем свойство к категории
            $this->propertyCollection[$categoryName][$propName]['CODE'] = $prop['CODE'];
            $this->propertyCollection[$categoryName][$propName]['VALUES'] = $this->addToPropertyCollectionList($prop['CODE']);
            return ['CODE' => $prop['CODE'], 'VALUE' => $valueID];
        }
        else
        {
            $result = $this->createProperty($propName, $propValue, $categoryName);
            return ['CODE' => $result['CODE'], 'VALUE' => $result['VALUE']];
        }
    }

    private function findOrCreateValue($propID, $propValue)
    {
        $propertyEnum = CIBlockPropertyEnum::GetList(Array(
            "DEF" => "DESC",
            "SORT" => "ASC"
        ) , Array(
            "IBLOCK_ID" => $this->IBLOCK_ID,
            "PROPERTY_ID" => $propID,
            'VALUE' => $propValue
        ));

        if ($enumField = $propertyEnum->GetNext())
        {
            return $enumField['ID'];
        }
        else
        {
            $ibpEnum = new CIBlockPropertyEnum;
            if ($propLID = $ibpEnum->Add(Array(
                'PROPERTY_ID' => $propID,
                'VALUE' => $propValue
            ))) return $propLID;
        }
    }

    private function createProperty($propName, $firstValue, $categoryName)
    {

        $resultProp = $this->createList($propName, $firstValue); //Создаем свойство
        CIBlockSectionPropertyLink::Add($this->sectId, $resultProp[0], $arLink = array(
            'IBLOCK_ID' => $this->IBLOCK_ID,
            'SMART_FILTER' => 'Y'
        )); // Привязываем свойство к категории
        $propertyEnum = CIBlockPropertyEnum::GetList(Array(
            "DEF" => "DESC",
            "SORT" => "ASC"
        ) , Array(
            "IBLOCK_ID" => $this->IBLOCK_ID,
            "PROPERTY_ID" => $resultProp[0],
            'VALUE' => $firstValue
        ));

        if ($enumField = $propertyEnum->GetNext())
        {
            $valueID = $enumField['ID'];
        }

        $this->propertyCollection[$categoryName][$propName]['CODE'] = $resultProp[1];

        $this->propertyCollection[$categoryName][$propName]['TYPE'] = 'L';

        $this->propertyCollection[$categoryName][$propName]['VALUES'] = $this->addToPropertyCollectionList($resultProp[1]);

        return ['CODE' => $resultProp[1], 'VALUE' => $valueID];

    }

    private function createList($propName, $firstValue)
    {
        $arFields = Array(
            "NAME" => $propName,
            "ACTIVE" => "Y",
            "SORT" => "500",
            "CODE" => $this->translit($propName) ,
            "PROPERTY_TYPE" => "L",
            "IBLOCK_ID" => $this->IBLOCK_ID,
            'SMART_FILTER' => 'Y'
        );

        $arFields["VALUES"][0] = Array(
            "VALUE" => $firstValue,
            "DEF" => "N",
            "SORT" => "500"
        );

        $ibp = new CIBlockProperty;
        $propID = $ibp->Add($arFields);

        return [$propID, $this->translit($propName) ];
    }

    private function checkMainInfo($item)
    {
        if (empty($item['Название']) || empty($item['Цена']) || empty($item['Ед'])) return false;
        else return true;
    }

    private function getProperties($categoryName)
    {

        if (!empty($this->propertyCollection[$categoryName])) return $this->propertyCollection[$categoryName];

        $rsSect = CIBlockSection::GetList(array() , array(
            'NAME' => $categoryName,
            'IBLOCK_ID' => $this->IBLOCK_ID
        ));
        print_r($categoryName);
        if ($arSect = $rsSect->GetNext())
        {
            $this->sectId = $arSect['ID'];
            foreach (CIBlockSectionPropertyLink::GetArray($arSect['IBLOCK_ID'], $arSect['ID']) as $PID => $arLink)
            {
                $res = CIBlockProperty::GetByID($PID);
                while ($resArr = $res->GetNext())
                {
                    $this->propertyCollection[$categoryName][$resArr['NAME']]['CODE'] = $resArr['CODE'];
                    if ($resArr['PROPERTY_TYPE'] === 'L')
                    {
                        $this->propertyCollection[$categoryName][$resArr['NAME']]['TYPE'] = 'L';
                        $this->propertyCollection[$categoryName][$resArr['NAME']]['VALUES'] = $this->addToPropertyCollectionList($resArr['CODE']);
                    }
                    else
                    {
                        $this->propertyCollection[$categoryName][$resArr['NAME']]['TYPE'] = $resArr['PROPERTY_TYPE'];
                    }
                }
            }
        }

    }

    private function addToPropertyCollectionList($propCode)
    {
        $propsList = [];
        $propertyEnums_q = CIBlockPropertyEnum::GetList(Array(
            "value" => "ASC",
            "SORT" => "ASC"
        ) , Array(
            "IBLOCK_ID" => $this->IBLOCK_ID,
            "CODE" => $propCode
        )); //ищем список значений свойства
        while ($propertyEnums = $propertyEnums_q->GetNext())
        {
            $propsList[$propertyEnums['VALUE']] = array(
                'CODE' => $propertyEnums['ID'],
                'VALUE' => $propertyEnums['VALUE']
            );
        }
        return $propsList;
    }

    private function prepareMainInfo()
    {
        global $USER;
        $data = CUser::GetList(($by = "ID") , ($order = "ASC") , array(
            'ID' => $USER->GetId()
        ));

        if ($arUser = $data->Fetch())
        {
            $this->sellerName = $arUser['WORK_COMPANY'];
        }
    }

    public function __construct($path, $limit = 0, $start = 0)
    {

        $this->itemsSheet = $this->prepareArrayItems($path, $limit, $start);

        $this->prepareMainInfo();

        $this->processItems();

    }
}
$feed = new processFeeds($inputFileName, 150);
