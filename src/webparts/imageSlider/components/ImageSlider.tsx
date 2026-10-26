
import * as React from "react";
import { useEffect } from "react";

import type { IImageSliderProps } from "./IImageSliderProps";
import { Swiper, SwiperSlide } from "swiper/react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";

import {
  EffectCoverflow,
  Autoplay,
  Navigation,
  Pagination,
} from "swiper/modules";

import "swiper/swiper-bundle.css";

import "./style.css";
 

interface IFolderItem {
  Name: string;
}

interface IFileItem {
  Name: string;
  ServerRelativeUrl: string;
}

const ImageSlider: React.FC<IImageSliderProps> = ({
  sp,
  selectedLibrary,
}) => {

  const [folders, setFolders] = React.useState<string[]>([]);
  const [selectedFolder, setSelectedFolder] = React.useState<string>("All");
  const [images, setImages] = React.useState<string[]>([]);
  const [libraryRoot, setLibraryRoot] = React.useState<string>("");
  const [swiperInstance, setSwiperInstance] = React.useState<any>(null);

  // -----------------------------
  // LIBRARY ROOT URL
  // -----------------------------
  const loadLibraryRoot = async (): Promise<void> => {
    if (!selectedLibrary) return;

    try {
      const lib = await sp.web.lists
        .getByTitle(selectedLibrary)
        .select("RootFolder/ServerRelativeUrl")
        .expand("RootFolder")();

      setLibraryRoot(lib.RootFolder.ServerRelativeUrl);
    } catch (err) {
      console.log(err);
    }
  };
 
  const loadFolders = async (): Promise<void> => {
  if (!selectedLibrary) return;

  try {
    const res: IFolderItem[] = await sp.web.lists
      .getByTitle(selectedLibrary)
      .rootFolder
      .folders
      .select("Name")();

    const folderNames = res
      .map((f) => f.Name)
      .filter(
        (name) =>
          name !== "Forms" &&
          !name.startsWith("_") &&
          !name.startsWith(".")
      );

    setFolders(folderNames);
    setSelectedFolder("All");

  } catch (err) {
    console.log(err);
  }
};
 

  const getImages = async (path: string): Promise<string[]> => {
    try {
      const files: IFileItem[] = await sp.web
        .getFolderByServerRelativePath(path)
        .files.select("Name", "ServerRelativeUrl")();

      return files
        .filter((f) =>
          /\.(jpg|jpeg|png|gif|webp)$/i.test(f.Name)
        )
        .map((f) => f.ServerRelativeUrl);

    } catch {
      return [];
    }
  };

  // -----------------------------
  // IMAGES FILTER
  // -----------------------------
  const loadImages = async (folder: string): Promise<void> => {
    if (!libraryRoot) return;

    let result: string[] = [];

    if (folder === "All") {

      result.push(...await getImages(libraryRoot));
      for (const f of folders) {
        result.push(...await getImages(`${libraryRoot}/${f}`));
      }

    } else {
      result.push(...await getImages(`${libraryRoot}/${folder}`));
    }

    setImages(result);
  };
 

  useEffect(() => {
    loadLibraryRoot();
  }, [selectedLibrary]);

  useEffect(() => {
    if (libraryRoot) {
      loadFolders();
    }
  }, [libraryRoot]);

  useEffect(() => {
    if (libraryRoot) {
      loadImages(selectedFolder);
    }
  }, [selectedFolder, folders, libraryRoot]);
useEffect(() => {
  if (swiperInstance && images.length > 0) {
    // wait a tick so slide DOM nodes actually exist
    const timer = setTimeout(() => {
      swiperInstance.update();
      if (swiperInstance.params.loop) {
        swiperInstance.loopDestroy();
        swiperInstance.loopCreate();
      }
      swiperInstance.slideToLoop(0, 0); // reset to a clean starting position
      swiperInstance.autoplay.stop();
      swiperInstance.autoplay.start();
    }, 50);
    return () => clearTimeout(timer);
  }
}, [images, swiperInstance]);
  // ----------------------------- Frontend ----------------------------- //
  return (
    <div className="imageSlider">

      {/* Folder Filter */}
      <div className="filterBox">
        <select
          value={selectedFolder}
          onChange={(e) => setSelectedFolder(e.target.value)}
        >
          <option value="All">All</option>

          {folders.map((folder, i) => (
            <option key={i} value={folder}>
              {folder}
            </option>
          ))}
        </select>
      </div>

      <div className="custom-prev"><svg
          xmlns="http://www.w3.org/2000/svg"
          width="22"
          height="22"
          viewBox="0 0 24 24"
          fill="none"
        >
          <path
            d="M14.5 6L8.5 12L14.5 18"
            stroke="currentColor"
            stroke-width="2.5"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg></div>
      <div className="custom-next"><svg
          xmlns="http://www.w3.org/2000/svg"
          width="22"
          height="22"
          viewBox="0 0 24 24"
          fill="none"
        >
          <path
            d="M9.5 6L15.5 12L9.5 18"
            stroke="currentColor" 
            stroke-width="2.5"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg></div>

      <Swiper
      onSwiper={setSwiperInstance}
        effect={"coverflow"}
        grabCursor={true}
        centeredSlides={true}
        observer={true}
        slidesPerView={2.8}
        spaceBetween={-190}
        loop={true}
        autoplay={{ delay: 5000, disableOnInteraction: false }}
        coverflowEffect={{
          rotate: 0,
          stretch: 0,
          depth: 100,
          modifier: 2.2, 
          scale: 0.85,
          slideShadows: false,
        }}
        pagination={{ clickable: true }}
        navigation={{
          prevEl: ".custom-prev",
          nextEl: ".custom-next",
        }} 
        modules={[EffectCoverflow, Autoplay, Pagination, Navigation]}
        className="customSwiper"
      >
        {images.map((img, i) => (
          <SwiperSlide key={i}> 
            <img src={img} alt={`slide-${i}`} /> 
          </SwiperSlide> 
        ))}
      </Swiper>

    </div>
  );
};

export default ImageSlider;