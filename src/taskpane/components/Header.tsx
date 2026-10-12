import * as React from "react";
import { Image, tokens, makeStyles } from "@fluentui/react-components";

export interface HeaderProps {
  title?: string;
  logo: string;
  message?: string;
}

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    paddingTop: "20px",
    paddingBottom: "10px",
    backgroundColor: tokens.colorNeutralBackground1, // Reines Weiß
  },
  logo: {
    width: "110px",
    height: "110px",
    objectFit: "contain",
  },
});

const Header: React.FC<HeaderProps> = ({ title, logo }) => {
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image className={styles.logo} src={logo} alt={title || "Logo"} />
    </section>
  );
};

export default Header;